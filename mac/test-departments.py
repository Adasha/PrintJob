# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-departments.ps1 - same checks, same messages, same order;
# see that script for the Departments sheet and the on-entry disregard stamp.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, addr, col_letters, sheet  # noqa: E402

DWS, DTBL = 'Departments', 'tblDepartments'
MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'

with open_copy() as wb:
    wb.run('SetQuiet', True)
    site = wb.call('SettingText', 'SITE_ID', 'SITE')

    section('The sheet', first=True)
    for h in ('DeptID', 'Name', 'Aliases', 'Free', 'Active', 'Notes', 'Match list', 'Dis Paper %', 'Dis Cons %',
              'Allowance', 'Allowance period'):
        check(wb.table_col(DWS, DTBL, h) > 0, f'column {h} exists')
    check(wb.column_hidden(DWS, wb.table_col(DWS, DTBL, 'Match list')), 'Match list is hidden')
    check(wb.prop(sheet(DWS), 'Visible') == '-1', 'sheet is visible')
    names = wb.sheet_names()
    i_dep, i_set, i_pap = names.index('Departments'), names.index('Settings'), names.index('Papers')
    check(i_pap < i_dep < i_set, 'tab order: Papers, Departments, Settings')
    hwm_cell = wb.wb_name('SET_DEPT_ID_HWM')
    check(wb.cell_format(*hwm_cell)['locked'] is True, 'DEPT_ID_HWM value cell is locked (Read-only)')
    check(int(wb.number(*hwm_cell)) == 0, 'DEPT_ID_HWM starts at 0')

    def dcell(row_no, h):
        return addr(wb.table_first_row(DWS, DTBL) + row_no, wb.table_col(DWS, DTBL, h))

    def dval(row_no, h):
        v = wb.value(DWS, dcell(row_no, h))
        return '' if v is None else str(v)

    def dset(row_no, h, v):
        wb.set_value(DWS, dcell(row_no, h), v)

    def set_dept(row_no, name, aliases, free, active):
        wb.call('UnlockSheet', sheet(DWS))
        dset(row_no, 'Name', name)
        if aliases is not None:
            dset(row_no, 'Aliases', aliases)
        if free:
            dset(row_no, 'Free', free)
        if active:
            dset(row_no, 'Active', active)

    section('Add row / typed rows')
    wb.activate(DWS)
    wb.run('AddCatalogRow', DTBL)
    id1 = dval(1, 'DeptID')
    check(re.match(rf'^{re.escape(site)}-DEP-0001$', id1) is not None, f'first ID is {id1}')
    check(dval(1, 'Active') == 'Yes', 'Active defaults to Yes')
    check(dval(1, 'Free') == 'No', 'Free defaults to No')
    set_dept(1, 'Fine Art', ' FA ; Art School;;', 'Yes', None)
    check(dval(1, 'Aliases') == 'FA;Art School', f"aliases tidied to '{dval(1, 'Aliases')}'")
    check(dval(1, 'Match list') == ';fine art;fa;art school;', f"match list is '{dval(1, 'Match list')}'")

    wb.run('AddCatalogRow', DTBL)
    set_dept(2, 'Textiles', '', None, None)
    wb.run('AddCatalogRow', DTBL)
    set_dept(3, 'Old Dept', '', 'Yes', 'No')
    # A row typed straight under the table (no Add row) still gets ID + defaults.
    wb.run('AddCatalogRow', DTBL)
    wb.call('UnlockSheet', sheet(DWS))
    dset(4, 'Name', 'Typed Dept')
    check(re.match(rf'^{re.escape(site)}-DEP-0004$', dval(4, 'DeptID')) is not None, 'typed row gets DEP-0004')
    check(dval(4, 'Active') == 'Yes' and dval(4, 'Free') == 'No', 'typed row gets Active=Yes, Free=No')
    check(int(wb.number(*hwm_cell)) == 4, 'counter advanced to 4')

    wb.run('SetQuiet', True)
    dset(4, 'Name', 'fa')
    log = wb.run('QuietLog')
    check('used more than once' in log, "a name that is another department's alias is warned about")
    check(dval(4, 'Name') == 'fa', '...and not blocked')
    dset(4, 'Name', 'Typed Dept')

    section('Job entry')
    wb.call('UnlockSheet', sheet(MAIN))

    def j(row_no, h):
        return addr(wb.table_first_row(MAIN, TBL) + row_no, wb.table_col(MAIN, TBL, h))

    def jset(row_no, h, v):
        wb.set_value(MAIN, j(row_no, h), v)

    def jnum(row_no, h):
        v = wb.value(MAIN, j(row_no, h))
        return 0.0 if v is None else float(v)

    def reset_row(row_no):
        jset(row_no, 'Disregard Paper', 'No')
        jset(row_no, 'Disregard Consumable', 'No')

    def dis(row_no):
        return f"{wb.value(MAIN, j(row_no, 'Disregard Paper'))}/{wb.value(MAIN, j(row_no, 'Disregard Consumable'))}"

    reset_row(1)
    gross = jnum(1, 'Gross Cost')
    check(gross > 0, f'row 1 has a cost to disregard ({gross})')
    check(jnum(1, 'Chargeable Cost') == gross, 'charged in full to begin with')

    jset(1, 'Student Name', ' fine ART ')
    check(dis(1) == 'Yes/Yes', f'name in odd case and padded: {dis(1)}')
    check(jnum(1, 'Gross Cost') == gross, 'Gross unchanged: still costed')
    check(jnum(1, 'Chargeable Cost') == 0, 'Chargeable is 0')
    check(abs(jnum(1, 'Disregarded') - gross) < 0.0001, 'Disregarded equals Gross')

    reset_row(2)
    jset(2, 'Student Name', 'FA')
    check(dis(2) == 'Yes/Yes', 'an alias matches')
    reset_row(3)
    jset(3, 'Student Name', 'Textiles')
    check(dis(3) == 'No/No', 'Free = No department changes nothing')
    reset_row(4)
    jset(4, 'Student Name', 'Old Dept')
    check(dis(4) == 'No/No', 'inactive department changes nothing')
    reset_row(5)
    jset(5, 'Student Name', 'Alex Student')
    check(dis(5) == 'No/No', 'an ordinary student changes nothing')

    # Hand override is never re-applied by an unrelated edit.
    jset(1, 'Disregard Paper', 'No')
    jset(1, 'Notes', 'edited')
    check(dis(1) == 'No/Yes', f'a hand-set value survives other edits: {dis(1)}')
    jset(1, 'Student No', 'X123')
    check(dis(1) == 'No/Yes', 'and editing Student No does not apply the rule')
    # ...but re-entering the department re-stamps.
    jset(1, 'Student Name', 'Textiles')
    check(dis(1) == 'No/Yes', 'switching to a non-free department leaves them')
    jset(1, 'Student Name', 'Fine Art')
    check(dis(1) == 'Yes/Yes', 'entering a free department again re-stamps')
    jset(1, 'Student Name', 'Alex Student')
    check(dis(1) == 'Yes/Yes', 'moving to a student leaves the values')

    # Renaming the department on the sheet: the old spelling no longer matches.
    reset_row(6)
    dset(1, 'Name', 'Fine Arts')
    jset(6, 'Student Name', 'Fine Art')
    check(dis(6) == 'No/No', 'after renaming the department, the old name no longer matches')
    jset(6, 'Student Name', 'Fine Arts')
    check(dis(6) == 'Yes/Yes', 'the new name does')

    section('Reporting')
    jset(1, 'Student Name', 'fine arts')
    jset(2, 'Student Name', 'FA')
    jset(3, 'Student Name', 'Textiles')
    jset(4, 'Student Name', 'Old Dept')
    jset(6, 'Student Name', 'Zed Student')     # was left on the department name by the section above
    wb.calculate_full_rebuild()
    heads = wb.texts('_Data', 'A9:BH9')[0]
    hdrs = {h: i + 1 for i, h in enumerate(heads) if h}
    check('Department' in hdrs, '_Data has a Department column')
    rows = int(wb.prop(xlmac.spill('_Data', 'A10'), 'Rows.Count'))
    dc = hdrs.get('Department', 1)
    block = wb.texts('_Data', f'A10:{col_letters(dc)}{9 + rows}')
    counts = {'Fine Arts': 0, 'Textiles': 0, 'Old Dept': 0, '': 0}
    bad_non_blank = 0
    for r in block:
        d = r[dc - 1]
        if d in counts:
            counts[d] += 1
        else:
            bad_non_blank += 1
    check(counts['Fine Arts'] == 2, f"two jobs are classed Fine Arts (the name and the alias) - got {counts['Fine Arts']}")
    check(counts['Textiles'] == 1, f"one is Textiles - got {counts['Textiles']}")
    check(counts['Old Dept'] == 1, f"an INACTIVE department still classifies its jobs - got {counts['Old Dept']}")
    check(bad_non_blank == 0, 'no other department name appears')
    check(counts[''] == rows - 4, f"every other row is blank (got {counts['']} of {rows})")

    REP = 'Reports'
    wb.activate(REP)

    def jobs():
        wb.calculate_full_rebuild()
        return int(wb.text(REP, 'I9'))

    every = jobs()
    wb.set_value(REP, 'B6', 'Fine Arts')
    check(jobs() == 2, 'Reports Department filter = Fine Arts shows 2 jobs')
    wb.set_value(REP, 'B6', 'Textiles')
    check(jobs() == 1, '...Textiles shows 1')
    wb.set_value(REP, 'B6', 'Old Dept')
    check(jobs() == 1, '...an inactive department can still be filtered')
    check(wb.validation_type(REP, 'B6') == 'validate list', 'B6 is a dropdown')
    wb.clear_contents(REP, 'B6')
    check(jobs() == every, 'cleared, back to all jobs')
    wb.set_value(REP, 'B6', 'Fine Arts')
    wb.call('ClearReportFilters', sheet(REP))
    check(wb.text(REP, 'B6') == '', 'Clear all filters clears the Department box')
    wb.set_value(REP, 'B6', 'Fine Arts')
    check(wb.call('HasActiveFilter', sheet(REP)) == 'True', 'a set Department filter counts as an active filter')
    wb.clear_contents(REP, 'B6')

    # The Reports breakdown (column AK) lists departments only.
    check(wb.cell_text(REP, 10, 37) == 'By department', 'Reports has a By department block')
    blk = wb.texts(REP, 'AK12:AL20')
    vals = wb.values(REP, 'AK12:AL20')
    names_b = [r[0] for r in blk if r[0]]
    jobs_col = {r[0]: int(float(v[1] or 0)) for r, v in zip(blk, vals) if r[0]}
    check(','.join(names_b) == 'Fine Arts,Old Dept,Textiles', f"breakdown keys: {','.join(names_b)}")
    check(jobs_col.get('Fine Arts') == 2 and jobs_col.get('Textiles') == 1, 'breakdown job counts are right')

    # The Summary box reconciles to the Summary totals.
    SUM = 'Summary'
    wb.activate(SUM)
    wb.calculate_full_rebuild()
    check(wb.text(SUM, 'I5') == 'By department', 'Summary has the By department box')
    box_t = wb.texts(SUM, 'I7:I21')
    box_v = wb.values(SUM, 'J7:M21')
    dn = []
    jobs_t = g_t = ds_t = ch_t = 0.0
    for k, ((n,), v) in enumerate(zip(box_t, box_v)):
        r = 7 + k
        if r <= 20 and n:
            dn.append(n)
        if r == 21 or n:
            jobs_t += float(v[0] or 0)
            g_t += float(v[1] or 0)
            ds_t += float(v[2] or 0)
            ch_t += float(v[3] or 0)
    check(','.join(dn) == 'Fine Arts,Old Dept,Textiles', f"box lists departments with jobs, sorted: {','.join(dn)}")
    check(wb.text(SUM, 'I21') == 'Everyone else', 'last line is Everyone else')
    check(abs(jobs_t - wb.number(SUM, 'B24')) < 0.001, f'jobs reconcile to the total ({jobs_t})')
    check(abs(g_t - wb.number(SUM, 'D24')) < 0.005, f'gross reconciles ({g_t})')
    check(abs(ds_t - wb.number(SUM, 'F24')) < 0.005, f'disregarded reconciles ({ds_t})')
    check(abs(ch_t - wb.number(SUM, 'H24')) < 0.005, f'chargeable reconciles ({ch_t})')
    check(wb.number(SUM, 'L7') > 0, 'Fine Arts shows a disregarded amount (free department)')

    wb.run('SetQuiet', False)

print()
if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('PASS')
