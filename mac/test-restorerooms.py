# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-restorerooms.ps1 - same checks, same messages, same order:
# back up a populated copy (A: Example Print Room + Annexe), restore it into a
# fresh copy (B) with no Annexe, and check the room is created once. The two
# copies run one after the other in the same Excel; B keeps A's folder, which
# holds the backup.
#
# Drives COPIES in a temp folder, never src/PrintJob.xlsm.

import glob
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, sheet  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

section('Back up copy A (Example Print Room + Annexe)', first=True)
a = open_copy()
orig_annexe_jobs, orig_dept = 0, ''
with a as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    # A department on the room, so the backup has one to carry.
    wb.call('UnlockSheet', sheet('Annexe'))
    wb.set_value('Annexe', wb.name_addr('Annexe', 'LOC_Dept'), 'Fine Art')
    wb.call('RelockSheet', sheet('Annexe'))
    orig_dept = wb.call('LocValue', sheet('Annexe'), 'LOC_Dept')
    orig_annexe_jobs = wb.table_row_count('Annexe', 'tblJobs_ANNEX')
    print(f"  Annexe: {orig_annexe_jobs} job row(s), department '{orig_dept}'")

    wb.run('BackupAll')
    wb.run('SetQuiet', False)

check(orig_annexe_jobs >= 1, f"the fixture's Annexe has job rows to back up (got {orig_annexe_jobs})")
annexe_csv = next((p for p in glob.glob(os.path.join(a.dir, 'PrintCosts-*.csv'))
                   if '-CATALOG-' not in p and 'ANNEX' in os.path.basename(p)), None)
check(annexe_csv is not None, 'Annexe job CSV found')
any_csv = next(iter(glob.glob(os.path.join(a.dir, 'PrintCosts-*-CATALOG-tblPrinters-*.csv'))), None)
check(any_csv is not None, 'a backup file to point Restore at found')

# The department is written beside the location name in the header block.
if annexe_csv:
    with open(annexe_csv, encoding='utf-8-sig') as f:
        hdr = [next(f, '').rstrip('\r\n') for _ in range(8)]
    check('Location department' in hdr[5] and orig_dept in hdr[5], f'job CSV header names the department (row 6: {hdr[5]})')

section('Restore into fresh copy B, which has no Annexe')
with open_copy(clean=False) as wb:
    wb.run('SetQuiet', True)

    def sheet_names():
        return [n.strip() for n in wb._do('return name of every worksheet of wb').split(',')]

    check('Annexe' not in sheet_names(), 'fresh copy has no Annexe sheet before the restore')

    if any_csv:
        wb.run('RestoreFromFileConfirmed', any_csv)
        print(wb.run('QuietLog'))
        time.sleep(0.3)

        check('Annexe' in sheet_names(), 'Annexe sheet was created by the restore')
        code = wb.call('LocValue', sheet('Annexe'), 'LOC_Code')
        check(code == 'ANNEX', f"created room has code ANNEX (got '{code}')")
        check(wb.call('LocValue', sheet('Annexe'), 'LOC_Name') == 'Annexe', 'created room has name Annexe')
        dept = wb.call('LocValue', sheet('Annexe'), 'LOC_Dept')
        check(dept == orig_dept, f"created room has department '{orig_dept}' (got '{dept}')")
        n = wb.table_row_count('Annexe', 'tblJobs_ANNEX')
        check(n == orig_annexe_jobs, f'Annexe job table holds {orig_annexe_jobs} row(s) (got {n})')

        actions = wb.table_column_texts('_Audit', 'tblAudit', wb.table_headers('_Audit', 'tblAudit')[2])
        check('Add print room' in actions, 'audit log records the room creation')
        check('Restore workbook' in actions, 'audit log records the restore')

        section('Restoring again creates no second room')
        wb.run('RestoreFromFileConfirmed', any_csv)
        time.sleep(0.3)
        count = sum(1 for nm in sheet_names() if nm.startswith('Annexe'))
        check(count == 1, f'still exactly one Annexe sheet after a second restore (got {count})')
        n = wb.table_row_count('Annexe', 'tblJobs_ANNEX')
        check(n == orig_annexe_jobs, f'Annexe still {orig_annexe_jobs} row(s) (got {n})')

    wb.run('SetQuiet', False)

print()
if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('PASS')
