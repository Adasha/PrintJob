# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-restorereuse.ps1 - same checks, same messages, same order;
# see that script for the older-workbook-with-the-sample's-code case. The
# older workbook (Old.xlsm) and the fresh one run one after the other in the
# same Excel; the fresh one keeps Old.xlsm's folder. Old.xlsm is saved by
# Excel itself, which is also what lets Excel read it back for the restore.
#
# Drives COPIES in a temp folder, never src/PrintJob.xlsm.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, sheet  # noqa: E402

WS = 'Example Print Room'


def restore_case(label, room_name, sheet_expected, code='EXAMPLE', first=False):
    section(label, first=first)
    a = open_copy(name='Old.xlsm')
    jobs_a, printers, code_a = 0, '', ''
    with a as wb:
        time.sleep(2)
        wb.run('SetQuiet', True)
        models = wb.table_column_texts('Printers', 'tblPrinters', 'Model')
        printers = f'{models[0]};{models[1]}'
        wb.call('UnlockSheet', sheet(WS))
        for nm, v in (('LOC_Name', room_name), ('LOC_Dept', 'Photography'),
                      ('LOC_Printers', printers), ('LOC_DefPrinter', models[0])):
            wb.set_value(WS, wb.name_addr(WS, nm), v)
        wb.call('RelockSheet', sheet(WS))
        if code != 'EXAMPLE':
            # An older workbook that shipped the sample as MAIN.
            wb.call('UnlockSheet', sheet(WS))
            wb.set_value(WS, wb.name_addr(WS, 'LOC_Code'), code)
            wb.call('RelockSheet', sheet(WS))
            wb.run('RefreshLocations')
        code_a = wb.text(WS, wb.name_addr(WS, 'LOC_Code'))
        jobs_a = wb.table_row_count(WS, f'tblJobs_{code}')
        check(jobs_a >= 1, f'old room has job rows ({jobs_a})')
        check(';' in printers and not printers.endswith(';'), f'old room permits two printers ({printers})')
        wb.save()
        wb.run('SetQuiet', False)
    old_path = os.path.join(a.dir, 'Old.xlsm')

    with open_copy(clean=False) as wb:
        time.sleep(2)
        wb.run('SetQuiet', True)
        wb.run('RestoreFromFileConfirmed', old_path)
        time.sleep(0.3)

        names = wb.sheet_names()
        rooms = sum(1 for s in names if wb.call('IsLocation', sheet(s)) == 'True')
        check(rooms == 1, f'exactly one print room remains ({rooms})')
        check(sheet_expected in names, f"the restored room's sheet is '{sheet_expected}' (sheets: {', '.join(names)})")
        if sheet_expected in names:
            got_code = wb.call('LocValue', sheet(sheet_expected), 'LOC_Code')
            check(got_code == code, f"code kept as {code} (got '{got_code}', A had '{code_a}')")
            check(wb.call('LocValue', sheet(sheet_expected), 'LOC_Name') == room_name, 'room name restored')
            check(wb.call('LocValue', sheet(sheet_expected), 'LOC_Dept') == 'Photography', 'department restored')
            got_printers = wb.call('LocValue', sheet(sheet_expected), 'LOC_Printers')
            check(got_printers == printers, f"permitted printers restored (got '{got_printers}')")
            n = wb.table_row_count(sheet_expected, f'tblJobs_{code}')
            check(n == jobs_a, f'job rows match the old room ({jobs_a}, got {n})')
        wb.run('SetQuiet', False)


restore_case('Sample room renamed in the older workbook', 'Photo Studio', 'Photo Studio', first=True)
restore_case('Sample room NOT renamed, only given its own printers', 'Example Print Room', 'Example Print Room')
restore_case('Older workbook that shipped the sample as MAIN', 'Example Print Room', 'Example Print Room', 'MAIN')

print()
if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('PASS')
