# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-backup.ps1 - same checks, same messages, same order: back up
# a populated copy (A), corrupt a second fresh copy (B), restore B from A's
# backup. The two copies are driven one after the other in the same Excel
# (the .ps1 starts two instances); B keeps A's folder, which holds the backup.
#
# Drives two COPIES in a temp folder, never src/PrintJob.xlsm.

import glob
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, sheet  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

section('Back up copy A', first=True)
a = open_copy()
with a as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    orig_printer_count = wb.table_row_count('Printers', 'tblPrinters')
    orig_paper_count = wb.table_row_count('Papers', 'tblPapers')
    orig_annexe_jobs = wb.table_row_count('Annexe', 'tblJobs_ANNEX')
    print(f'  original counts: Printers={orig_printer_count} Papers={orig_paper_count} Annexe jobs={orig_annexe_jobs}')

    wb.run('BackupAll')
    print(wb.run('QuietLog'))

    wb.run('SetQuiet', False)

catalog_csvs = glob.glob(os.path.join(a.dir, 'PrintCosts-*-CATALOG-*.csv'))
check(len(catalog_csvs) == 9, f'9 catalogue CSVs produced (got {len(catalog_csvs)})')
printers_csv = next((p for p in catalog_csvs if '-CATALOG-tblPrinters-' in os.path.basename(p)), None)
check(printers_csv is not None, 'tblPrinters backup CSV found')

location_csvs = [p for p in glob.glob(os.path.join(a.dir, 'PrintCosts-*.csv')) if '-CATALOG-' not in p]
check(len(location_csvs) == 2, f'2 location job CSVs produced (got {len(location_csvs)})')

section("Corrupt copy B, then restore it from copy A's backup")
with open_copy(clean=False) as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    def empty(ws_name, table):
        wb.call('UnlockSheet', sheet(ws_name))
        while wb.table_row_count(ws_name, table) > 0:
            wb.delete_table_row(ws_name, table, 1)
        wb.call('RelockSheet', sheet(ws_name))

    empty('Printers', 'tblPrinters')
    empty('Annexe', 'tblJobs_ANNEX')

    check(wb.table_row_count('Printers', 'tblPrinters') == 0, 'tblPrinters is empty before restore')
    check(wb.table_row_count('Annexe', 'tblJobs_ANNEX') == 0, 'Annexe job table is empty before restore')

    if printers_csv:
        wb.run('RestoreFromFileConfirmed', printers_csv)
        print(wb.run('QuietLog'))
        time.sleep(0.3)   # the .ps1's settling pause, kept

        n = wb.table_row_count('Printers', 'tblPrinters')
        check(n == orig_printer_count, f'tblPrinters restored to {orig_printer_count} rows (got {n})')
        n = wb.table_row_count('Papers', 'tblPapers')
        check(n == orig_paper_count,
              f"tblPapers unchanged at {orig_paper_count} rows (got {n}) - was not part of this restore's file set")
        n = wb.table_row_count('Annexe', 'tblJobs_ANNEX')
        check(n == orig_annexe_jobs, f'Annexe job table restored to {orig_annexe_jobs} row(s) (got {n})')

        actions = wb.table_column_texts('_Audit', 'tblAudit', wb.table_headers('_Audit', 'tblAudit')[2])
        last_action = actions[-1] if actions else ''
        check(last_action == 'Restore workbook', f"audit log entry recorded (got '{last_action}')")

        section('Restoring again is idempotent (no duplicate rows)')
        wb.run('RestoreFromFileConfirmed', printers_csv)
        n = wb.table_row_count('Printers', 'tblPrinters')
        check(n == orig_printer_count, f'tblPrinters still {orig_printer_count} rows after a second restore (got {n})')

    wb.run('SetQuiet', False)

print()
if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('PASS')
