# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-multipass.ps1 - same checks, same messages, same order; see
# that script for multi-pass (per-colour) costing. Every macro that takes the
# sheet or the job table gets it as a modScript token.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, addr, sheet, table  # noqa: E402

CWS, CTBL = 'Consumables', 'tblColours'
PWS, PTBL = 'Printers', 'tblPrinters'
MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'


def near(a, b):
    return abs(float(a or 0) - b) < 0.0001


with open_copy() as wb:
    wb.run('SetQuiet', True)
    site = wb.call('SettingText', 'SITE_ID', 'SITE')

    section('Consumables sheet and a multi-pass printer', first=True)
    for h in ('ColourID', 'Consumable type', 'Colour', 'Cost per m2', 'Active'):
        check(wb.table_col(CWS, CTBL, h) > 0, f'colours column {h} exists')
    names = wb.sheet_names()
    check(names.index('Papers') < names.index('Consumables') < names.index('Settings'), 'tab order: Papers, Consumables ... Settings')
    hwm = wb.wb_name('SET_COLOUR_ID_HWM')
    check(int(wb.number(*hwm)) == 0 and wb.cell_format(*hwm)['locked'], 'COLOUR_ID_HWM starts at 0 and is read-only')

    def ccell(row_no, h):
        return addr(wb.table_first_row(CWS, CTBL) + row_no, wb.table_col(CWS, CTBL, h))

    def set_colour(row_no, name, typ, rate):
        wb.call('UnlockSheet', sheet(CWS))
        wb.set_value(CWS, ccell(row_no, 'Colour'), name)
        wb.set_value(CWS, ccell(row_no, 'Consumable type'), typ)
        wb.set_value(CWS, ccell(row_no, 'Cost per m2'), rate)

    wb.activate(CWS)
    for name, typ, rate in (('Black', 'Risograph', 0.30), ('Blue', 'Risograph', 0.40),
                            ('Fluoro Orange', 'Risograph', 0.55), ('Cyan', 'Ink', 0.10)):
        wb.run('AddCatalogRow', CTBL)
        set_colour(wb.table_row_count(CWS, CTBL), name, typ, rate)
    check(wb.text(CWS, ccell(1, 'ColourID')) == f'{site}-CLR-0001', f'first colour ID is {site}-CLR-0001')
    check(wb.text(CWS, ccell(1, 'Active')) == 'Yes', 'Active defaults to Yes')
    check(int(wb.number(*hwm)) == 4, 'counter advanced to 4')

    wb.activate(PWS)
    wb.run('AddCatalogRow', PTBL)
    prow = wb.table_last_row(PWS, PTBL)
    wb.call('UnlockSheet', sheet(PWS))
    for h, v in (('Model', 'Test RISO'), ('Consumable type', 'Risograph'), ('Cost per m2', 0.99),
                 ('Colour mode', 'multi-pass'), ('Template cost', 3), ('Max sheet size', 'A1'), ('Active', 'Yes')):
        wb.set_value(PWS, addr(prow, wb.table_col(PWS, PTBL, h)), v if isinstance(v, str) else float(v))

    loc = wb.name_addr(MAIN, 'LOC_Printers')
    wb.call('UnlockSheet', sheet(MAIN))
    wb.set_value(MAIN, loc, wb.text(MAIN, loc) + ';Test RISO')

    first = wb.table_first_row(MAIN, TBL)

    def j(row_no, h):
        return addr(first + row_no, wb.table_col(MAIN, TBL, h))

    def jv(row_no, h):
        v = wb.value(MAIN, j(row_no, h))
        return '' if v is None else v

    def js(row_no, h):
        v = jv(row_no, h)
        return v if isinstance(v, str) else (f'{v:g}' if isinstance(v, float) else str(v))

    def w(row_no, h, v):
        """A write from here: unprotected first; Worksheet_Change runs as for a person typing."""
        wb.call('UnlockSheet', sheet(MAIN))
        wb.set_value(MAIN, j(row_no, h), v if isinstance(v, str) else float(v))

    def pick(row_no):
        wb.activate(MAIN)
        wb.select(MAIN, j(row_no, 'Date/Time'))

    def level(row_no):
        return wb.row_outline_level(MAIN, first + row_no)

    def hidden(row_no):
        return wb.row_hidden(MAIN, first + row_no)

    def rows():
        return wb.table_row_count(MAIN, TBL)

    def log():
        return wb.run('QuietLog')

    sample_rows = rows()

    section('A job on the multi-pass printer')
    wb.activate(MAIN)
    wb.call('AddPrintJob', sheet(MAIN))
    job = rows()
    job_id = js(job, 'Job ID')
    check(js(job, 'Row Type') == 'Job', 'a new job is a Job row')
    check(level(job) == 1, 'at outline level 1')
    w(job, 'Student Name', 'Pass Tester')
    w(job, 'Technician', js(1, 'Technician'))
    w(job, 'Printer', 'Test RISO')
    w(job, 'Paper Stock', 'Gloss 200gsm SRA3 sheet')
    w(job, 'Qty', 100)
    check(js(job, 'S_PrinterID') != '', 'the job is stamped with the printer')
    check(js(job, 'S_ConsRate') == '', 'a multi-pass job carries no flat consumable rate')
    st = js(job, 'Status')
    check('At least one colour pass is required' in st, f"Status asks for a pass: '{st}'")
    check(js(job, 'Passes') == '', 'Passes is blank with none')
    paper, area = float(jv(job, 'Paper Cost') or 0), float(jv(job, 'Area m2') or 0)
    check(near(area, 14.4), f'area is 14.4 m2 (got {area})')
    check(near(paper, 31.0), f'paper cost is 31.00 (got {paper})')
    check(float(jv(job, 'Consumable Cost') or 0) == 0, 'consumable cost is 0 with no passes')

    section('Add pass')
    pick(job)
    wb.run('SetQuiet', True)
    for _ in range(3):
        wb.call('AddPass', sheet(MAIN))
    check(rows() == job + 3, f'three pass rows added beneath the job (rows now {rows()})')
    for i in (1, 2, 3):
        r = job + i
        check(js(r, 'Row Type') == 'Pass', f'row {r} is a Pass row')
        check(js(r, 'Parent') == job_id, f'pass {i} has Parent {job_id}')
        check(int(float(jv(r, 'Pass') or 0)) == i, f'pass number {i}')
        check(level(r) == 2, f'pass {i} is at outline level 2')
        check(float(jv(r, 'S_SetupCost') or 0) == 3, f"pass {i} carries the printer's Template cost")
        check(js(r, 'Job ID') == '', f'pass {i} has no Job ID of its own')
    check(level(job) == 1, 'the job stays at level 1')
    check(int(float(jv(job, 'Passes') or 0)) == 3, 'Passes counts 3')
    check(float(jv(job, 'Set-up Cost') or 0) == 9, 'Set-up Cost rolls up to 9.00')
    st = js(job + 1, 'Status')
    check('Colour required' in st, f"an unfilled pass asks for a colour: '{st}'")

    section('Colours and costing')
    w(job + 1, 'Colour', 'Black')
    w(job + 2, 'Colour', 'Blue')
    w(job + 3, 'Colour', 'Fluoro Orange')
    check(float(jv(job + 1, 'S_ConsRate') or 0) == 0.30, 'Black stamps 0.30')
    check(js(job + 1, 'S_ColourID') == f'{site}-CLR-0001', 'and its ColourID')
    check(near(jv(job + 1, 'Consumable Cost'), 4.32), 'pass 1 ink = 14.4 x 0.30 = 4.32')
    check(near(jv(job + 2, 'Consumable Cost'), 5.76), 'pass 2 ink = 5.76')
    check(near(jv(job + 3, 'Consumable Cost'), 7.92), 'pass 3 ink = 7.92')
    check(near(jv(job, 'Consumable Cost'), 18.0), 'the job sums its ink: 18.00')
    check(near(jv(job, 'Set-up Cost'), 9.0), 'and its set-up: 9.00')
    check(near(jv(job, 'Gross Cost'), 58.0), 'Gross = 31 + 18 + 9 = 58.00')
    check(near(jv(job, 'Chargeable Cost'), 58.0), 'Chargeable is the full 58.00')
    check(js(job + 1, 'Gross Cost') == '', 'a pass row has no gross of its own')
    st = js(job, 'Status')
    check(st == 'OK', f"the job is OK: '{st}'")
    check(js(job + 2, 'Status') == 'OK', 'a coloured pass is OK')

    w(job, 'Disregard Consumable', 'Yes')
    check(near(jv(job, 'Chargeable Cost'), 31.0), 'Disregard Consumable drops ink AND set-up: chargeable 31.00')
    check(near(jv(job, 'Disregarded'), 27.0), 'Disregarded = 27.00')
    w(job, 'Disregard Consumable', 'No')

    # A later price change never reaches the stamped rate (snapshot rule, decision 23).
    wb.activate(CWS)
    wb.call('UnlockSheet', sheet(CWS))
    wb.set_value(CWS, ccell(1, 'Cost per m2'), 0.90)
    check(near(jv(job + 1, 'Consumable Cost'), 4.32), 'editing the colour price leaves the pass at its stamped rate')
    wb.call('UnlockSheet', sheet(CWS))
    wb.set_value(CWS, ccell(1, 'Cost per m2'), 0.30)

    section('Problems versus notices')
    w(job + 2, 'Colour', 'Mystery')
    st = js(job + 2, 'Status')
    check(st.startswith('Notice:'), f"an unknown colour shows a notice: '{st}'")
    check(js(job + 2, 'H_Issues') == '', '...and is not a problem')
    check(near(jv(job + 2, 'Consumable Cost'), 5.76), '...and still counts at its stamped rate')
    wb.run('SetQuiet', True)
    wb.call('CheckSheet', sheet(MAIN))
    lg = log()
    check('Notices (not counted as problems)' in lg, 'Check sheet reports notices apart')
    check('Pass ' not in lg.split('Notices (not counted')[0], '...and none of the passes is listed as a problem')
    w(job + 2, 'Colour', 'Cyan')
    st = js(job + 2, 'Status')
    check('different consumable type' in st, f"a colour of another type is a problem: '{st}'")
    w(job + 2, 'Colour', 'Blue')
    check(js(job + 2, 'Status') == 'OK', 'a Blue pass is OK again')
    # The Pass column is locked for a person; here it is written with events off to prove the check.
    wb.set_enable_events(False)
    w(job + 3, 'Pass', 7)
    wb.set_enable_events(True)
    check('out of sequence' in js(job + 3, 'H_Issues'), 'a pass number out of sequence is a problem')
    wb.set_enable_events(False)
    w(job + 3, 'Pass', 3)
    wb.set_enable_events(True)
    st = js(job + 3, 'Status')
    check(st == 'OK', f"put back, the pass is OK ('{st}')")

    section('The job keeps its printer; single-pass jobs refuse a pass')
    wb.run('SetQuiet', True)
    w(job, 'Printer', 'Xerox Versant 180')
    check(js(job, 'Printer') == 'Test RISO', 'changing the printer of a job with passes is reverted')
    check('cannot be changed' in log(), '...with a message')
    pick(1)
    wb.run('SetQuiet', True)
    wb.call('AddPass', sheet(MAIN))
    check('Select a multi-pass job, or one of its passes, first.' in log(), 'Add pass on a single-pass job shows the short message')
    check(rows() == job + 3, '...and adds nothing')

    section('Collapse and expand')
    pick(job + 2)                     # a pass selected: acts on its job
    wb.call('TogglePasses', sheet(MAIN))
    check(hidden(job + 1) and hidden(job + 2) and hidden(job + 3), 'Toggle passes collapses the job from one of its passes')
    check(not hidden(job), '...the job row stays visible')
    wb.call('TogglePasses', sheet(MAIN))
    check(not (hidden(job + 1) or hidden(job + 3)), 'a second click expands')
    wb.set_row_hidden(MAIN, first + job + 1, True)     # a mix
    wb.call('ToggleAllPasses', sheet(MAIN))
    check(not hidden(job + 1), 'Toggle all with a mix expands everything first')
    wb.call('ToggleAllPasses', sheet(MAIN))
    check(hidden(job + 1) and hidden(job + 3), 'all expanded -> collapse all')
    wb.call('ToggleAllPasses', sheet(MAIN))
    check(not hidden(job + 2), 'all collapsed -> expand all')

    section('Remove a pass: renumbering')
    wb.call('DeletePass', sheet(MAIN), table(TBL), job, job + 2, 'test')
    check(rows() == job + 2, 'one row fewer')
    check(int(float(jv(job + 1, 'Pass') or 0)) == 1 and int(float(jv(job + 2, 'Pass') or 0)) == 2, 'the rest are numbered 1, 2')
    check(js(job + 2, 'Colour') == 'Fluoro Orange', '...and it was the middle one that went')
    check(int(float(jv(job, 'Passes') or 0)) == 2, 'Passes now 2')
    check(near(jv(job, 'Set-up Cost'), 6.0), 'Set-up Cost now 6.00')
    check(level(job + 1) == 2 and level(job + 2) == 2 and level(job) == 1, 'outline levels still line up')

    section("Passes, Pass, Unit and Area m2 are locked, as calculated columns are")
    check(wb.prop(f'@@range:{MAIN}!{j(job, "Passes")}', 'Locked') == 'True', "the job's Passes cell is locked")
    check(wb.prop(f'@@range:{MAIN}!{j(job + 1, "Pass")}', 'Locked') == 'True', "a pass's Pass cell is locked")
    check(wb.prop(f'@@range:{MAIN}!{j(job + 1, "Colour")}', 'Locked') == 'False', "a pass's Colour cell is still open for typing")
    # Unit and Area m2: the template's column formats left rows added after
    # the sample unlocked (0.11.5).
    for r, what in ((job, "the job's"), (job + 1, "a pass's")):
        for h in ('Unit', 'Area m2'):
            check(wb.prop(f'@@range:{MAIN}!{j(r, h)}', 'Locked') == 'True', f'{what} {h} cell is locked')
    qty_text = wb.text(MAIN, j(job, 'Qty'))
    check(qty_text == '100', f"a whole Qty shows without a trailing point (got '{qty_text}')")

    section('Repeat job copies the group')
    pick(job + 1)
    wb.call('RepeatJob', sheet(MAIN))
    job2 = job + 3
    check(rows() == job + 5, f'repeat added a job and two passes (rows {rows()})')
    id2 = js(job2, 'Job ID')
    check(id2 != '' and id2 != job_id, f'the copy has its own Job ID ({id2})')
    check(js(job2 + 1, 'Parent') == id2 and js(job2 + 2, 'Parent') == id2, 'its passes belong to it')
    check(js(job2 + 1, 'Colour') == 'Black' and js(job2 + 2, 'Colour') == 'Fluoro Orange', 'with the same colours')
    check(level(job2) == 1 and level(job2 + 1) == 2 and level(job2 + 2) == 2, 'and the right outline levels')
    check(int(float(jv(job2, 'Passes') or 0)) == 2, 'Passes 2 on the copy')

    section('Counts and consolidation')
    n_jobs = sample_rows + 2
    b13 = int(wb.number(MAIN, 'B13'))
    check(b13 == n_jobs, f'Print jobs counts jobs, not passes ({b13} of {n_jobs})')
    spill_rows = int(wb.prop(xlmac.spill('_Data', 'A10'), 'Rows.Count'))
    check(spill_rows == n_jobs, f'the consolidated range holds jobs only ({spill_rows} of {n_jobs})')

    section('Delete a job takes its passes')
    before = rows()
    wb.call('DeleteJobGroupAt', sheet(MAIN), table(TBL), job2, 'test')
    check(rows() == before - 3, 'all three rows removed')
    check(level(job + 1) == 2 and level(job) == 1, 'the first group is undisturbed')
    wb.call('DeleteJobGroupAt', sheet(MAIN), table(TBL), job, 'test')
    check(rows() == sample_rows, 'back to the sample rows')

if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('multi-pass checks complete')
