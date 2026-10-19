# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-exportnames.ps1 - same checks, same messages, same order;
# see that script for the include-student-names choice it guards. Exports
# within one minute share a filename, so each file is copied aside straight
# after it is written, as the .ps1 does.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import csv
import glob
import io
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr, sheet  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

MAIN, MTBL = 'Example Print Room', 'tblJobs_EXAMPLE'
ANNEX, ATBL = 'Annexe', 'tblJobs_ANNEX'

ctx = open_copy()
with ctx as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    def set_cell(ws, row, col, v):
        wb.set_value(ws, addr(row, col), v if isinstance(v, str) else float(v))
        time.sleep(0.3)

    def cc(ws, table, name):
        return wb.table_col(ws, table, name)

    def find_row(ws, table, job_id):
        first = wb.table_first_row(ws, table) + 1
        ids = wb.table_column_texts(ws, table, 'Job ID')
        return first + ids.index(job_id) if job_id in ids else 0

    last_minute = [None]

    def export_main(name):
        """Export Main and return the file. The .ps1 copies each export aside
        because exports in the same minute share a filename. Here the file has
        to stay exactly where and as Excel wrote it - its sandbox refuses to
        read a copy, or even a rename ("Excel cannot access") - so instead a
        second export waits for the next minute, and gets a name of its own.
        name is the .ps1's label for the file, kept for the reader."""
        while time.strftime('%Y%m%d-%H%M') == last_minute[0]:
            time.sleep(1)
        last_minute[0] = time.strftime('%Y%m%d-%H%M')
        before = set(glob.glob(os.path.join(ctx.dir, 'PrintCosts-*-EXAMPLE-*.csv')))
        wb.activate(MAIN)
        wb.run('btnExport')
        new = sorted(set(glob.glob(os.path.join(ctx.dir, 'PrintCosts-*-EXAMPLE-*.csv'))) - before, key=os.path.getmtime)
        if not new:
            raise RuntimeError(f'export for {name} wrote no new file')
        return new[-1]

    def import_into(ws, path):
        rows = wb.call('ReadImportRows', path)
        wb.call('ApplyImportConfirmed', sheet(ws), rows)
        wb.calculate_full_rebuild()
        time.sleep(0.3)

    def read_csv(path):
        with open(path, encoding='utf-8-sig') as f:
            lines = f.read().splitlines()
        recs = list(csv.DictReader(io.StringIO('\n'.join(lines[8:]))))
        return lines, recs

    def export_status():
        return wb.text(MAIN, wb.name_addr(MAIN, 'LOC_Export'))

    # ------------------------------------------------ a job with a name
    section('Setup: one Main job with a student name and number', first=True)
    wb.activate(MAIN)
    wb.run('btnAddPrintJob')
    r = wb.table_last_row(MAIN, MTBL)
    m_id, m_name, m_no = cc(MAIN, MTBL, 'Job ID'), cc(MAIN, MTBL, 'Student Name'), cc(MAIN, MTBL, 'Student No')
    set_cell(MAIN, r, cc(MAIN, MTBL, 'Paper Stock'), 'Canvas 914mm roll')
    set_cell(MAIN, r, cc(MAIN, MTBL, 'Qty'), 2)
    set_cell(MAIN, r, m_name, 'Ada Lovelace')
    set_cell(MAIN, r, m_no, 'S1234')
    job_id = wb.cell_text(MAIN, r, m_id)
    check(len(job_id) > 0, f'job {job_id} created')

    # ---------------------------------------------- default: names in
    section('Quiet default: names included')
    with_names = export_main('with-names.csv')
    lines, recs = read_csv(with_names)
    check(not re.search('Student names', lines[7]), 'header block carries no omitted marker')
    rec = next((x for x in recs if x.get('Job ID') == job_id), {})
    check(rec.get('Student Name') == 'Ada Lovelace', f"Student Name in file ({rec.get('Student Name')})")
    check(rec.get('Student No') == 'S1234', 'Student No in file')
    st = export_status()
    check(st.startswith('Exported'), f'location stamped as exported ({st})')

    # ------------------------------------------------- names omitted
    section('Names omitted')
    set_cell(MAIN, r, cc(MAIN, MTBL, 'Qty'), 3)
    wb.call('RefreshExportStatus', sheet(MAIN))
    before = export_status()
    check(before.startswith('CHANGED'), f'status shows CHANGED after an edit ({before})')
    wb.run('SetExportNames', False)
    no_names = export_main('no-names.csv')
    wb.run('SetExportNames', True)
    lines, recs = read_csv(no_names)
    check(re.search('Student names,Omitted', lines[7]) is not None, f'header block marks names omitted ({lines[7]})')
    check(lines[8].startswith('Job ID,'), 'header row still on line 9')
    rec = next((x for x in recs if x.get('Job ID') == job_id), {})
    check(not rec.get('Student Name'), 'Student Name blank')
    check(rec.get('Student No') == 'S1234', 'Student No kept')
    for need in ('S_UnitCost', 'S_ConsRate', 'S_SchemaVer', 'Chargeable Cost', 'Notes'):
        check(need in rec, f'column {need} still present')
    check(export_status() == before, 'no-names export did not change the export status')

    # ------------------------------------------------------- imports
    section('Importing the no-names file')
    set_cell(MAIN, r, m_name, 'Changed Name')
    import_into(MAIN, no_names)
    r = find_row(MAIN, MTBL, job_id)
    check(wb.cell_text(MAIN, r, m_name) == 'Changed Name', 'overwrite kept the name already on the sheet')
    check(wb.cell_text(MAIN, r, m_no) == 'S1234', 'Student No still S1234')

    import_into(ANNEX, no_names)
    ar = find_row(ANNEX, ATBL, job_id)
    check(ar > 0, 'job appended to Annexe')
    if ar > 0:
        check(wb.cell_text(ANNEX, ar, cc(ANNEX, ATBL, 'Student Name')) == '', 'appended row has a blank name')
        check(wb.cell_text(ANNEX, ar, cc(ANNEX, ATBL, 'Student No')) == 'S1234', 'appended row has the Student No')

    section('Importing the with-names file still overwrites the name')
    import_into(MAIN, with_names)
    r = find_row(MAIN, MTBL, job_id)
    check(wb.cell_text(MAIN, r, m_name) == 'Ada Lovelace', 'name restored from the with-names file')

    wb.run('SetQuiet', False)

finish()
