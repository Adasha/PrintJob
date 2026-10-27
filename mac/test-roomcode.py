# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-roomcode.ps1 - same checks, same messages, same order.
# A print room code is used once (0.11.4): Add print room's prompt refuses a
# code another room has (RoomUsingCode, with FreeCode's suggestion), and
# Refresh Locations lets a room already registered keep its code when a new
# room repeats it. CreatePrintRoom is called directly, so there is no InputBox
# to answer.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, table, check, section, finish  # noqa: E402

FIRST, CLASH = 'Photography', 'Photo Lab'

with open_copy() as wb:
    wb.run('SetQuiet', True)
    reg_ws = next((s for s in wb.sheet_names() if 'tblLocations' in wb.table_names(s)), None)

    def reg_row(sheet_name):
        """The registry row for a sheet, as {header: text}, or None."""
        heads = wb.table_headers(reg_ws, 'tblLocations')
        cols = {h: wb.table_column_texts(reg_ws, 'tblLocations', h) for h in heads}
        for i, s in enumerate(cols['SheetName']):
            if s == sheet_name:
                return {h: cols[h][i] for h in heads}
        return None

    # ------------------------------------------------------------- prompt
    section('The prompt knows which codes are taken', first=True)
    got = wb.call('RoomUsingCode', 'example')
    check(got == 'Example Print Room', f"'example' is the sample room's code (got {got!r})")
    got = wb.call('RoomUsingCode', 'NOPE')
    check(got == '', f"'NOPE' is nobody's code (got {got!r})")
    got = wb.call('FreeCode', 'nope')
    check(got == 'NOPE', f'an unused code is suggested as typed (got {got!r})')

    ws = wb.call('CreatePrintRoom', FIRST, 'Media Services', 'PHOTO')
    check(ws.startswith('@@h:'), f"CreatePrintRoom made '{FIRST}'")
    ids = [wb.call('NextJobId', ws, table('tblJobs_PHOTO')) for _ in range(2)]
    check(ids == ['UNI-PHOTO-00001', 'UNI-PHOTO-00002'], f"'{FIRST}' handed out two job IDs ({ids})")
    got = wb.call('RoomUsingCode', 'photo')
    check(got == FIRST, f"'photo' is now '{FIRST}''s code (got {got!r})")
    got = wb.call('FreeCode', 'PHOTO')
    check(got == 'PHOTO2', f'the suggestion after a clash is PHOTO2 (got {got!r})')

    # ------------------------------------------------------------ refresh
    section("A new room repeating a code does not take it from the room that has it")
    ws2 = wb.call('CreatePrintRoom', CLASH, '', 'PHOTO')
    check(ws2.startswith('@@h:'), f"CreatePrintRoom made '{CLASH}'")
    names = wb.sheet_names()
    check(CLASH in names and FIRST in names and names.index(CLASH) < names.index(FIRST),
          f"'{CLASH}' sits in front of '{FIRST}', the order that went wrong (sheets: {', '.join(names)})")
    got = wb.call('LocValue', ws, 'LOC_Code')
    check(got == 'PHOTO', f"'{FIRST}' keeps PHOTO (got {got!r})")
    got = wb.call('LocValue', ws2, 'LOC_Code')
    check(got == 'PHOTO2', f"'{CLASH}' is given PHOTO2 (got {got!r})")
    row = reg_row(FIRST) or {}
    check(row.get('Code') == 'PHOTO' and row.get('Job ID HWM') == '2',
          f"the registry still has '{FIRST}' on PHOTO with its high-water mark of 2 ({row})")
    row = reg_row(CLASH) or {}
    check(row.get('Code') == 'PHOTO2', f"the registry has '{CLASH}' on PHOTO2 ({row})")
    check('tblJobs_PHOTO' in wb.table_names(FIRST) and 'tblJobs_PHOTO2' in wb.table_names(CLASH),
          'each room\'s job table is named after its own code')
    got = wb.call('NextJobId', ws2, table('tblJobs_PHOTO2'))
    check(got == 'UNI-PHOTO2-00001', f"'{CLASH}''s first job ID is UNI-PHOTO2-00001 (got {got!r})")
    got = wb.call('NextJobId', ws, table('tblJobs_PHOTO'))
    check(got == 'UNI-PHOTO-00003', f"'{FIRST}' carries on from UNI-PHOTO-00003 (got {got!r})")

    wb.run('SetQuiet', False)

finish()
