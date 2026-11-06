# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-excelcheck.ps1 - same checks, same messages, same order.
# An Excel missing LET and the other newer functions (0.12.1). Writing the
# workbook's formulas into one stores them broken, so the commands that write
# them refuse first (modUtils.ExcelCanRun). No Excel with the fault is to hand:
# the real check is tried against an unknown function, and the refusals are
# reached through PretendMissingExcelFunctions.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, sheet, check, section, finish  # noqa: E402

ROOM = 'Example Print Room'
TBL = 'tblJobs_EXAMPLE'

with open_copy() as wb:
    wb.run('SetQuiet', True)

    section('The check', first=True)
    got = wb.call('MissingExcelFunctions')
    check(got == '', f'this Excel has every function (missing: {got!r})')
    got = wb.call('ExcelKnowsFormula', '=LET(x,1,x)')
    check(got == 'True', f'a known function is known (got {got!r})')
    got = wb.call('ExcelKnowsFormula', '=NOSUCHFUNCTION(1)')
    check(got == 'False', f'an unknown function is not (got {got!r})')
    wb.call('WarnIfExcelFunctionsMissing')
    got = wb.run('QuietLog')
    check(got == '', f'no warning on opening (got {got!r})')

    extra = wb.call('CreatePrintRoom', 'Second Room', '', 'SECOND')
    check(extra.startswith('@@h:'), 'a second room to work on')
    data_before = wb.formula('_Data', 'A10')
    sheets_before = wb.sheet_names()
    wb.run('SetQuiet', True)

    section('Pretending LET and VSTACK are missing')
    wb.call('PretendMissingExcelFunctions', 'LET, VSTACK')
    got = wb.call('MissingExcelFunctions')
    check(got == 'LET, VSTACK', f'the check reports them (got {got!r})')
    wb.call('WarnIfExcelFunctionsMissing')
    got = wb.run('QuietLog')
    check('LET, VSTACK' in got and 'will not work properly' in got, 'the workbook warns on opening')
    wb.run('SetQuiet', True)

    def refused(what, macro, *args):
        wb.run('SetQuiet', True)
        wb.call(macro, *args)
        log = wb.run('QuietLog')
        check(f'{what} did not run' in log and 'LET, VSTACK' in log, f'{what} refuses, naming them')

    refused('Refresh Locations', 'RefreshLocations')
    check(wb.formula('_Data', 'A10') == data_before, "_Data's formula is untouched")
    check(TBL in wb.table_names(ROOM), 'job tables were not parked')
    refused('Add print room', 'CreatePrintRoom', 'Refused Room', '', 'NOPE')
    refused('Rename print room', 'ChangeRoomName', sheet('Second Room'), 'Renamed Room')
    refused('Remove print room', 'DeletePrintRoom', sheet('Second Room'))
    refused('Restore workbook', 'RestoreFromFileConfirmed', '/nonexistent.xlsm')
    got = wb.sheet_names()
    check(got == sheets_before, f'no sheet came or went, or changed name ({got})')

    section('Formulas are not rewritten while it is missing')
    wb.unprotect(ROOM)
    wb.set_enable_events(False)
    wb.clear_contents(ROOM, 'B12')
    wb.call('EnsurePrintersDisplay', sheet(ROOM))
    check(wb.formula(ROOM, 'B12') == '', 'the room header is left alone')

    section('With the functions back')
    wb.call('PretendMissingExcelFunctions', '')
    got = wb.call('MissingExcelFunctions')
    check(got == '', f'the check reports nothing (got {got!r})')
    wb.call('EnsurePrintersDisplay', sheet(ROOM))
    got = wb.text(ROOM, 'B12')
    check(got.startswith('3 printers'), f'the room header is written (got {got!r})')
    wb.set_enable_events(True)
    wb.run('SetQuiet', True)
    wb.call('RefreshLocations')
    got = wb.run('QuietLog')
    check('2 print rooms registered' in got, f'Refresh Locations runs ({got[:60]!r})')

    wb.run('SetQuiet', False)

finish()
