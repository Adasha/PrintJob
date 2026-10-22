# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-restoreexample.ps1 - same checks, same messages, same order;
# see that script for when Restore workbook removes the shipped sample room.
# The copies run one after the other in the same Excel; each later one keeps
# the backup copy's folder.
#
# Drives COPIES in a temp folder, never src/PrintJob.xlsm.

import glob
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, sheet, addr  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402


def restore_fresh(csv_path, prep=None):
    """Restore csv_path into a fresh copy; the sheet names afterwards."""
    with open_copy(clean=False) as wb:
        wb.run('SetQuiet', True)
        if prep:
            prep(wb)
        wb.run('RestoreFromFileConfirmed', csv_path)
        print(wb.run('QuietLog'))
        time.sleep(0.3)
        names = wb.sheet_names()
        wb.run('SetQuiet', False)
        return names


section('Back up a copy holding Example Print Room + Annexe', first=True)
a = open_copy()
with a as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)
    wb.run('BackupAll')
    wb.run('SetQuiet', False)

full = next(iter(glob.glob(os.path.join(a.dir, 'PrintCosts-*-CATALOG-tblPrinters-*.csv'))), None)
check(full is not None, 'a backup file to point Restore at found')
main_csv = next((p for p in glob.glob(os.path.join(a.dir, 'PrintCosts-*.csv'))
                 if '-CATALOG-' not in p and 'EXAMPLE' in os.path.basename(p)), None)
check(main_csv is not None, 'EXAMPLE job CSV found')

if full:
    section('2. Backup holds EXAMPLE and Annexe: Example Print Room stays')
    names = restore_fresh(full)
    check('Example Print Room' in names, 'Example Print Room kept when the backup includes the sample room')
    check('Annexe' in names, 'Annexe created')

    # Take the sample room out of the backup set for the remaining cases.
    if main_csv:
        os.remove(main_csv)

    section('1. Backup holds Annexe only: Example Print Room is removed')
    names = restore_fresh(full)
    check('Annexe' in names, 'Annexe created')
    check('Example Print Room' not in names, f"Example Print Room removed (sheets: {', '.join(names)})")

    section('3. Example room holds a real job: it stays')

    def add_real_job(wb):
        ws, t = 'Example Print Room', 'tblJobs_EXAMPLE'
        wb.call('UnlockSheet', sheet(ws))
        idx = int(wb.prop(xlmac.table(t), 'ListRows.Add.Index'))
        wb.set_value(ws, addr(wb.table_first_row(ws, t) + idx, wb.table_col(ws, t, 'Job ID')), 'UNI-EXAMPLE-00013')
        wb.call('RelockSheet', sheet(ws))

    names = restore_fresh(full, add_real_job)
    check('Annexe' in names, 'Annexe created')
    check('Example Print Room' in names, 'Example Print Room kept because it holds a job that is not a sample')

print()
if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('PASS')
