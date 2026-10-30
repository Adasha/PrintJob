# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac twin of test-picker.ps1 - same checks, same messages, same order.
# The Select printers... picker sheet (_Picker) placed its checkboxes and
# OK/Cancel from the cells' positions BEFORE setting its column widths, and
# the buttons are free-floating, so they stayed wherever the sheet's previous
# widths had put them (found in the 2026-10-30 Mac screenshot session; fixed
# in 0.11.5, modPicker.Show). Here the sheet starts with deliberately wrong
# widths, then every control must sit where the final layout puts it.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, sheet, shape  # noqa: E402

MAIN, PICK = 'Example Print Room', '_Picker'
FIRST_ROW = 6   # modPicker.FIRST_ROW

with open_copy() as wb:
    wb.run('SetQuiet', True)

    # Stale widths, as a previous use of the sheet could leave them.
    wb.call('UnlockSheet', sheet(PICK))
    for c in (1, 2, 3, 4):
        wb.set_column_width(PICK, c, 20)
    wb.call('RelockSheet', sheet(PICK))

    section('Select printers... places its controls after its widths', first=True)
    wb.activate(MAIN)
    wb.call('PickPrinters', sheet(MAIN))
    lefts = wb.column_lefts(PICK, 5)
    b_left, c_left, c_right = lefts[1], lefts[2], lefts[3]

    n = 0
    while True:
        r = FIRST_ROW + n
        try:
            left = float(wb.prop(shape(PICK, f'pick_{r}'), 'Left'))
            width = float(wb.prop(shape(PICK, f'pick_{r}'), 'Width'))
        except Exception:
            break
        check(c_left <= left and left + width <= c_right,
              f'the row {r} tick sits inside the Include column ({left:.1f}-{left + width:.1f} in {c_left:.1f}-{c_right:.1f})')
        n += 1
    check(n > 0, f'the picker lists printers (got {n})')

    btns = {b['on_action'].split('!')[-1]: b for b in wb.buttons(PICK)}
    ok, cancel = btns.get('btnPickerOK'), btns.get('btnPickerCancel')
    check(ok is not None and cancel is not None, 'OK and Cancel are drawn')
    if ok and cancel:
        top, _ = wb.row_top_height(PICK, FIRST_ROW + n + 1)
        check(abs(float(ok['top']) - top) < 1 and abs(float(cancel['top']) - top) < 1,
              f'OK and Cancel sit one row below the list (tops {ok["top"]}, {cancel["top"]}; row {top})')
        check(abs(float(ok['left']) - b_left) < 1, f'OK starts at the Option column (left {ok["left"]}, column {b_left:.1f})')
        check(float(ok['left']) + float(ok['width']) <= float(cancel['left']),
              f'Cancel is to the right of OK, not over it (OK ends {float(ok["left"]) + float(ok["width"]):.1f}, Cancel starts {cancel["left"]})')
        right = float(cancel['left']) + float(cancel['width'])
        check(abs(right - c_right) < 1, f'Cancel ends at the Include column\'s right edge ({right:.1f}, column {c_right:.1f})')

    wb.run('btnPickerCancel')
    wb.run('SetQuiet', False)

finish()
