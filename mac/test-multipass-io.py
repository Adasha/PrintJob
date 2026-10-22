# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-multipass-io.ps1 - same checks, same messages, same order;
# see that script for Export / Import / Re-stamp / Backup of colour passes.
# Differences:
#
#   - The .ps1 writes its variant files (two passes, an orphan pass, an old
#     format, an unknown colour) from scratch. Excel's sandbox will not read a
#     file it did not write, so Excel saves each one (xlmac.write_csv, from a
#     helper module in a throwaway Builder.xlsm copy).
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import csv
import glob
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import (open_copy, check, section, addr, sheet, table, builder_helper,  # noqa: E402
                   WRITECSV_HELPER, write_csv)

MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'
CWS, CTBL = 'Consumables', 'tblColours'
PWS, PTBL = 'Printers', 'tblPrinters'


def near(a, b):
    return abs(float(a or 0) - b) < 0.0001


ctx = open_copy()
with builder_helper(WRITECSV_HELPER) as helper, ctx as wb:   # helper first: the copy stays the active workbook
    wb.run('SetQuiet', True)
    site = wb.call('SettingText', 'SITE_ID', 'SITE')
    first = wb.table_first_row(MAIN, TBL)

    def j(row_no, h):
        return addr(first + row_no, wb.table_col(MAIN, TBL, h))

    def jv(row_no, h):
        v = wb.value(MAIN, j(row_no, h))
        return '' if v is None else v

    def js(row_no, h):
        v = jv(row_no, h)
        return v if isinstance(v, str) else f'{v:g}'

    def ji(row_no, h):
        v = jv(row_no, h)
        return int(float(v)) if v != '' else 0

    def w(row_no, h, v):
        wb.call('UnlockSheet', sheet(MAIN))
        wb.set_value(MAIN, j(row_no, h), v if isinstance(v, str) else float(v))

    def pick(row_no):
        wb.activate(MAIN)
        wb.select(MAIN, j(row_no, 'Date/Time'))

    def level(row_no):
        return wb.row_outline_level(MAIN, first + row_no)

    def rows():
        return wb.table_row_count(MAIN, TBL)

    def latest_csv():
        found = sorted(glob.glob(os.path.join(ctx.dir, 'PrintCosts-*-EXAMPLE-*.csv')), key=os.path.getmtime)
        return found[-1] if found else None

    def import_file(path):
        wb.call('ApplyImportConfirmed', sheet(MAIN), wb.call('ReadImportRows', path))

    def job_row(job_id):
        ids = wb.table_column_texts(MAIN, TBL, 'Job ID')
        return ids.index(job_id) + 1 if job_id in ids else 0

    def variant(name, lines):
        """A file Excel will read: Excel itself saves it, as modExport does."""
        dst = os.path.join(ctx.dir, name)
        write_csv(helper, dst, list(csv.reader(lines)))
        return dst

    # ---- setup: colours, a RISO printer, one job with three passes ----------------
    wb.activate(CWS)
    for name, rate in (('Black', 0.30), ('Blue', 0.40), ('Fluoro Orange', 0.55)):
        wb.run('AddCatalogRow', CTBL)
        n = wb.table_row_count(CWS, CTBL)
        wb.call('UnlockSheet', sheet(CWS))
        crow = wb.table_first_row(CWS, CTBL) + n
        wb.set_value(CWS, addr(crow, wb.table_col(CWS, CTBL, 'Colour')), name)
        wb.set_value(CWS, addr(crow, wb.table_col(CWS, CTBL, 'Consumable type')), 'Risograph')
        wb.set_value(CWS, addr(crow, wb.table_col(CWS, CTBL, 'Cost per m2')), rate)
    wb.activate(PWS)
    wb.run('AddCatalogRow', PTBL)
    prow = wb.table_last_row(PWS, PTBL)
    wb.call('UnlockSheet', sheet(PWS))
    for h, v in (('Model', 'Test RISO'), ('Consumable type', 'Risograph'), ('Colour mode', 'multi-pass'),
                 ('Template cost', 3.0), ('Max sheet size', 'A1'), ('Active', 'Yes')):
        wb.set_value(PWS, addr(prow, wb.table_col(PWS, PTBL, h)), v)
    wb.call('UnlockSheet', sheet(MAIN))
    loc = wb.name_addr(MAIN, 'LOC_Printers')
    wb.set_value(MAIN, loc, wb.text(MAIN, loc) + ';Test RISO')
    sample = rows()

    def new_riso_job(student, colours):
        wb.activate(MAIN)
        wb.call('AddPrintJob', sheet(MAIN))
        jr = rows()
        w(jr, 'Student Name', student)
        w(jr, 'Technician', js(1, 'Technician'))
        w(jr, 'Printer', 'Test RISO')
        w(jr, 'Paper Stock', 'Gloss 200gsm SRA3 sheet')
        w(jr, 'Qty', 100)
        for c in colours:
            pick(jr)
            wb.call('AddPass', sheet(MAIN))
            w(rows(), 'Colour', c)
        return jr

    job = new_riso_job('Round Trip', ['Black', 'Blue', 'Fluoro Orange'])
    job_id = js(job, 'Job ID')
    gross0, cons0, setup0 = float(jv(job, 'Gross Cost') or 0), float(jv(job, 'Consumable Cost') or 0), float(jv(job, 'Set-up Cost') or 0)
    check(near(cons0, 18.0) and near(setup0, 9.0) and near(gross0, 58.0), f'setup job costs 18 + 9 + 31 = 58 (got {cons0} / {setup0} / {gross0})')

    section('A pass can be added to a job that is not the last row')
    job2 = new_riso_job('Second', ['Black'])
    pick(job)
    wb.call('AddPass', sheet(MAIN))
    new_pass = job + 4
    check(js(new_pass, 'Row Type') == 'Pass' and js(new_pass, 'Parent') == job_id, 'the 4th pass lands directly under the first job')
    check(level(new_pass) == 2, '...at level 2')
    check(js(job + 5, 'Row Type') == 'Job', '...and the next job moved down intact')
    check(js(new_pass, 'Consumable Cost') in ('0', ''), 'its calculated columns filled (no colour yet: ink 0)')
    check('Colour required' in js(new_pass, 'Status'), '...Status asks for its colour (the Status formula filled in)')
    wb.call('DeletePass', sheet(MAIN), table(TBL), job, new_pass, 'test')
    check(ji(job, 'Passes') == 3, 'removed again: Passes is 3')
    wb.call('DeleteJobGroupAt', sheet(MAIN), table(TBL), job + 4, 'test')     # the second job and its pass
    check(rows() == job + 3, 'second job removed; back to one group')

    section('Export')
    wb.activate(MAIN)
    wb.run('btnExport')
    src_csv = latest_csv()
    check(src_csv is not None, 'a CSV was produced')
    with open(src_csv, encoding='utf-8-sig') as f:
        lines = f.read().splitlines()
    hdr = next(csv.reader([lines[8]]))
    print('  header: ' + ' | '.join(hdr))
    check(hdr[:4] == ['Job ID', 'Row Type', 'Parent', 'Pass'], 'Row Type, Parent and Pass follow Job ID')
    check(hdr.index('Passes') == hdr.index('Printer') + 1 and hdr.index('Set-up Cost') == hdr.index('Consumable Cost') + 1,
          'Passes follows Printer; Set-up Cost follows Consumable Cost')
    check('S_SetupCost' in hdr and 'S_ColourID' in hdr and 'Colour' in hdr, 'Colour, S_SetupCost and S_ColourID are in the file')
    data = list(csv.DictReader(io.StringIO('\n'.join(lines[8:]))))
    pass_rows = [d for d in data if d['Row Type'] == 'Pass']
    check(len(pass_rows) == 3, f"the file holds the job's 3 pass rows (got {len(pass_rows)})")
    check(sum(1 for d in pass_rows if d['Parent'] == job_id) == 3, '...each naming its job in Parent')
    check(sum(1 for d in pass_rows if d['Job ID'] != '') == 0, '...with no Job ID of their own')
    check(sum(1 for d in pass_rows if d['Student Name'] != '' or d['Qty'] != '') == 0, '...and none of the job fields filled in')
    check(sum(1 for d in data if d['Row Type'] == 'Job') == sample + 1, 'every other row is a Job row')

    # The variants are made now, while the export is the latest file.
    seen = 0
    kept = list(lines[:9])
    for ln in lines[9:]:
        if ln.startswith(',Pass,'):
            seen += 1
            if seen == 3:
                continue
        kept.append(ln)
    two = variant('two-passes.csv', kept)
    orphan = variant('orphan.csv', lines[:9] + [',Pass,UNI-EXAMPLE-99999,1,,,,,,,Black,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,'])
    old_cols = ['Job ID', 'Date/Time', 'Student Name', 'Student No', 'Technician', 'Printer', 'Paper Stock', 'Unit', 'Qty',
                'Print Width mm', 'Sheet size', 'Disregard Paper', 'Disregard Consumable', 'Paid', 'Notes', 'S_PrinterID',
                'S_StockID', 'S_TechID', 'S_Measure', 'S_UnitCost', 'S_StockWidth_mm', 'S_SheetHeight_mm', 'S_ConsRate',
                'S_StampedAt', 'S_StampedBy', 'S_SchemaVer']
    old_rows = [{c: d.get(c, '') for c in old_cols} for d in data if d['Job ID'] == job_id]
    old_rows[0]['Notes'] = 'edited in old file'
    buf = io.StringIO()
    wr = csv.DictWriter(buf, fieldnames=old_cols, quoting=csv.QUOTE_ALL, lineterminator='\n')
    wr.writeheader()
    wr.writerows(old_rows)
    old_file = variant('old-format.csv', lines[:8] + buf.getvalue().splitlines())
    ghost = variant('ghost.csv', lines[:9] + [ln.replace(',Black,', ',Ghost Pink,') for ln in lines[9:]])

    section('Delete the job, import the file: job and passes come back')
    wb.call('DeleteJobGroupAt', sheet(MAIN), table(TBL), job, 'test')
    check(rows() == sample, 'job and its passes deleted')
    rset = wb.call('ReadImportRows', src_csv)
    rc = int(wb.prop(rset, 'Count'))
    check(rc == sample + 1 + 3, f'ReadImportRows returns every row, passes included ({rc})')
    wb.call('ApplyImportConfirmed', sheet(MAIN), rset)
    wb.calculate_full_rebuild()
    back = rows()
    check(back == sample + 4, f'four rows came back (now {back})')
    jr = job_row(job_id)
    check(jr > 0, 'the job is back')
    check(js(jr + 1, 'Row Type') == 'Pass' and ji(jr + 1, 'Pass') == 1 and js(jr + 1, 'Colour') == 'Black', 'pass 1 is Black')
    check(ji(jr + 2, 'Pass') == 2 and js(jr + 2, 'Colour') == 'Blue' and js(jr + 3, 'Colour') == 'Fluoro Orange',
          'passes 2 and 3 are Blue and Fluoro Orange')
    check(level(jr) == 1 and level(jr + 1) == 2 and level(jr + 3) == 2, 'outline levels are right')
    check(near(jv(jr, 'Consumable Cost'), cons0) and near(jv(jr, 'Set-up Cost'), setup0) and near(jv(jr, 'Gross Cost'), gross0),
          'costs match the original exactly (from the stamped rates, not recosted)')
    check(ji(jr, 'Passes') == 3, 'Passes is 3')

    section('Importing over an existing job replaces its passes as one block')
    import_file(src_csv)
    check(rows() == sample + 4, 'the same file again: still four rows, passes not doubled')
    import_file(two)
    check(rows() == sample + 3, 'a file with two passes leaves the job with two')
    check(ji(job_row(job_id), 'Passes') == 2, '...Passes 2')
    before = rows()
    import_file(orphan)
    check(rows() == before, 'a pass whose job is not in the file is skipped')
    import_file(src_csv)   # restore the 3-pass state

    section('A file from before 0.11.0 (no Row Type) leaves passes alone')
    import_file(old_file)
    jr = job_row(job_id)
    check(js(jr, 'Notes') == 'edited in old file', 'the old-format job row was imported')
    check(ji(jr, 'Passes') == 3 and js(jr + 1, 'Row Type') == 'Pass', 'and its passes were not touched')

    section('An unknown colour imports, keeps its cost, shows a notice')
    import_file(ghost)
    wb.calculate_full_rebuild()
    jr = job_row(job_id)
    check(js(jr + 1, 'Colour') == 'Ghost Pink', 'the colour came in as written')
    st = js(jr + 1, 'Status')
    check(st.startswith('Notice:'), f"it shows the notice: '{st}'")
    check(near(jv(jr, 'Consumable Cost'), cons0), 'and still counts at the stamped rate')

    section('Re-stamp prices')
    import_file(src_csv)   # back to known colours
    jr = job_row(job_id)
    wb.call('UnlockSheet', sheet(CWS))
    wb.set_value(CWS, addr(wb.table_first_row(CWS, CTBL) + 1, wb.table_col(CWS, CTBL, 'Cost per m2')), 0.50)   # Black 0.30 -> 0.50
    wb.call('UnlockSheet', sheet(PWS))
    wb.set_value(PWS, addr(prow, wb.table_col(PWS, PTBL, 'Template cost')), 4.0)                             # set-up 3 -> 4
    check(near(jv(jr + 1, 'S_ConsRate'), 0.30), 'before: the pass still has its stamped 0.30')
    wb.call('ReStampAllConfirmed', 1)
    wb.calculate_full_rebuild()
    check(near(jv(jr + 1, 'S_ConsRate'), 0.50), 'after: Black re-stamped to 0.50')
    check(near(jv(jr + 2, 'S_ConsRate'), 0.40), 'Blue unchanged at 0.40')
    check(near(jv(jr + 1, 'S_SetupCost'), 4), 'set-up cost re-stamped from the printer: 4')
    check(near(jv(jr, 'Set-up Cost'), 12.0), 'the job rolls up 3 x 4 = 12.00')
    # An unknown colour is left alone (decision 31).
    w(jr + 3, 'Colour', 'Not Defined Here')
    rate_before = float(jv(jr + 3, 'S_ConsRate') or 0)
    wb.call('ReStampAllConfirmed', 1)
    check(near(jv(jr + 3, 'S_ConsRate'), rate_before), 're-stamp leaves a pass with an unknown colour at its stamped rate')

    section('Backup includes the colours table')
    # Backup writes the jobs file under the same minute-stamped name as the export above; clear it so Excel does not ask to overwrite.
    for p in glob.glob(os.path.join(ctx.dir, 'PrintCosts-*-EXAMPLE-*.csv')):
        os.remove(p)
    wb.run('btnBackupWorkbook')
    cat = glob.glob(os.path.join(ctx.dir, 'PrintCosts-*-CATALOG-tblColours-*.csv'))
    check(len(cat) == 1, 'the colours table has its own backup CSV')
    printers_csv = glob.glob(os.path.join(ctx.dir, 'PrintCosts-*-CATALOG-tblPrinters-*.csv'))
    ptext = open(printers_csv[0], encoding='utf-8-sig').read() if printers_csv else ''
    check('Colour mode' in ptext and 'Template cost' in ptext, 'and the printers backup carries Colour mode and Template cost')

if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('multi-pass import/export checks complete')
