# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-roomname.ps1 - same checks, same messages, same order.
# An apostrophe in a print room's name (0.12.1): the LOC_ names are rebuilt from
# the sheet name on every refresh, and an unescaped apostrophe made that raise
# 1004. CreatePrintRoom and ChangeRoomName are called directly, so there is no
# InputBox to answer.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, sheet, check, section, finish  # noqa: E402

ROOM, RENAMED = "Dean's Room", "Bob's Lab"

with open_copy() as wb:
    def clean(log, what):
        check('went wrong' not in log, f'{what} raised no error ({log[:70]!r})')

    section('Adding a room with an apostrophe', first=True)
    wb.run('SetQuiet', True)
    ws = wb.call('CreatePrintRoom', ROOM, '', 'DEANS')
    clean(wb.run('QuietLog'), 'Add print room')
    check(ROOM in wb.sheet_names(), f"the tab is called {ROOM}")
    got = wb.name_addr(ROOM, 'LOC_Columns')
    check(got == 'AX11', f'LOC_Columns is on the room (got {got!r})')
    got = wb.name_addr(ROOM, 'LOC_DefTech')
    check(got == 'B4', f'LOC_DefTech is on the room (got {got!r})')
    got = wb.name_addr(ROOM, 'LOC_Export')
    check(got == 'AX8', f'LOC_Export is on the room (got {got!r})')
    got = wb.text(ROOM, 'B12')
    check(got.startswith('0 printers'), f'Printers at this location reads (got {got!r})')

    section('Refreshing')
    wb.run('SetQuiet', True)
    wb.call('RefreshLocations')
    log = wb.run('QuietLog')
    clean(log, 'Refresh Locations')
    check('2 print rooms registered' in log, 'both rooms are registered')
    check(wb.formula('_Data', 'A10') != '', '_Data has its formula')

    section('Renaming')
    wb.run('SetQuiet', True)
    wb.call('ChangeRoomName', sheet(ROOM), RENAMED)
    clean(wb.run('QuietLog'), 'Rename print room')
    check(RENAMED in wb.sheet_names(), f'the tab follows to {RENAMED}')
    got = wb.name_addr(RENAMED, 'LOC_DefTech')
    check(got == 'B4', f'LOC_DefTech is still on the room (got {got!r})')

    section('An apostrophe at either end')
    wb.run('SetQuiet', True)
    wb.call('CreatePrintRoom', "'Quoted'", '', 'QUOTED')
    clean(wb.run('QuietLog'), 'Add print room')
    check('Quoted' in wb.sheet_names(), 'the tab drops them')

    section('Removing')
    wb.run('SetQuiet', True)
    wb.call('DeletePrintRoom', sheet(RENAMED))
    clean(wb.run('QuietLog'), 'Remove print room')
    check(RENAMED not in wb.sheet_names(), 'the room has gone')

    wb.run('SetQuiet', False)

finish()
