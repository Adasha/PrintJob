# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac only - no .ps1 counterpart, because the bug it guards is Excel for Mac's.
#
# Excel for Mac 16.114: a workbook created while ScreenUpdating is False saves
# as an EMPTY file with no error, and the next close of the workbook that made
# it deletes that workbook's file from disk. Export, Export report and Backup
# all built their file in a scratch workbook inside AppOff, so a user
# exporting from a workbook in an ordinary folder got "exported", an empty
# CSV, and then lost the workbook. Fixed by modUtils.NewScratchBook /
# CloseScratchBook (screen updating on for the scratch book's whole life) and
# FileWritten (an empty file is not a success).
#
# Each writer runs into the workbook's own folder as Excel reports it - the
# case that failed - then the copy is closed and must still be on disk.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, sheet, WB_NAME  # noqa: E402

WRITERS = [
    ('Export', 'PrintCosts-UNI-EXAMPLE-*.csv'),
    ('Export report', 'PrintCosts-Report-*.xlsx'),
    ('Backup workbook', 'PrintCosts-*-CATALOG-*.csv'),
]


def run_writer(wb, label):
    if label == 'Export':
        wb.call('ExportLocation', sheet('Example Print Room'))
    elif label == 'Export report':
        wb.run('SetExportNames', False)
        wb.activate('Reports')
        wb.run('btnExportReport')
    else:
        wb.run('BackupAll')


for i, (label, pattern) in enumerate(WRITERS):
    section(f'{label}, into the workbook\'s own folder', first=(i == 0))
    ctx = open_copy()
    with ctx as wb:
        wb.run('SetQuiet', True)
        folder = wb.call('ExportFolder')
        run_writer(wb, label)
        log = wb.run('QuietLog')
        files = glob.glob(os.path.join(ctx.dir, pattern))
        check(len(files) > 0, f'{label} wrote a file into {folder}')
        empty = [os.path.basename(f) for f in files if os.path.getsize(f) == 0]
        check(files and not empty, f'every file it wrote has content (empty: {empty or "none"})')
        check('did not produce a file' not in log and 'ERROR' not in log,
              f'{label} reported success (log: {log.strip().splitlines()[0] if log.strip() else "empty"})')
    check(os.path.exists(os.path.join(ctx.dir, WB_NAME)), 'the workbook is still on disk after it is closed')

finish()
