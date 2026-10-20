# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-reports.ps1 - same checks, same messages, same order; see
# that script for Summary and Reports end to end (criteria, filters, layout,
# dropdowns, sorting, the minimum-columns view, Paid, the Summary toggle and
# Export report). Differences:
#
#   - Each exported .xlsx is read back by opening it in the same Excel (there
#     is no second Excel to start), as test-subsidy.py does; its page setup
#     and colours come through AppleScript's page setup object, interior
#     object and font object - except the paper size, which AppleScript
#     does not have, so it is read from the .xlsx's own sheet XML.
#   - Validation types come back as Excel's enum names ('validate list' is
#     COM's 3, a date one names 'date').
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import glob
import os
import re
import sys
import time
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import (open_copy, check, section, sheet, spill, Workbook, open_in_excel, osa,  # noqa: E402
                   ExcelError, col_letters, SUM_HDR_ROW, SUM_FIRST_ROW, SUM_TOT_ROW)
from fixture_annexe import add_annexe_fixture  # noqa: E402

REP, SUM, MAIN = 'Reports', 'Summary', 'Example Print Room'

ctx = open_copy()
with ctx as wb:
    add_annexe_fixture(wb)
    wb.calculate_full_rebuild()
    time.sleep(3)
    wb.run('SetQuiet', True)
    c = REP

    hdr10 = wb.texts(REP, 'H10:AL10')[0]

    def reports_col(header):
        """Position within the results table (1 = column H), 0 if absent."""
        return hdr10.index(header) + 1 if header in hdr10 else 0

    def t(a1):
        return wb.text(REP, a1)

    def jobs():
        return t('I9')

    def spill_block(ws, a1, text=False):
        ref = wb.prop(spill(ws, a1), 'Address').replace('$', '')
        return (wb.texts if text else wb.values)(ws, ref)

    section('Reports: hidden Job ID correlation column', first=True)
    job_id_col = reports_col('Job ID')
    check(job_id_col != 0, 'no Job ID header found on row 10')
    letter = col_letters(7 + job_id_col)
    check(wb.column_hidden(REP, 7 + job_id_col), f'column {letter} (Job ID) should be hidden')
    sp = spill_block(REP, 'H11')
    print(f'  results spill: {len(sp)} rows x {len(sp[0])} cols')
    sample_id = str(sp[0][job_id_col - 1])
    print(f"  {letter}16 (Job ID) sample value: '{sample_id}'")
    check('-EXAMPLE-' in sample_id or '-ANNEX-' in sample_id, 'hidden Job ID column does not look like a Job ID')
    print(f'  OK: Job ID is present as a hidden column ({letter})')

    section('Summary (Location x Printer x Paper Stock)')
    print(f"  Summary table formula length: {len(wb.formula(SUM, f'A{SUM_FIRST_ROW}'))} chars")
    try:
        ssp = spill_block(SUM, f'A{SUM_FIRST_ROW}')
        print(f'  spill: {len(ssp)} rows x {len(ssp[0])} cols')
    except ExcelError:
        print('  NO SPILL: ' + wb.text(SUM, f'A{SUM_FIRST_ROW}'))
    tot = wb.texts(SUM, f'A{SUM_TOT_ROW}:H{SUM_TOT_ROW}')[0]
    print(f'  totals: jobs={tot[1]} gross={tot[3]} disregarded={tot[5]} chargeable={tot[7]}')
    hdr_row = wb.texts(SUM, f'A{SUM_HDR_ROW}:M{SUM_HDR_ROW}')[0]
    print('  headers: ' + ' | '.join(hdr_row))
    check(hdr_row[1] == 'Printer', f"Summary column B header is 'Printer' (got '{hdr_row[1]}')")
    print('  OK: Printer is now part of the Summary key (Location, Printer, Paper stock, ...)')
    for r in wb.texts(SUM, 'A10:M15'):
        v = [r[i - 1] for i in (1, 2, 3, 5, 6, 7, 11, 13)]
        if v[0] != '':
            print('   ' + ' | '.join(v))

    section('Reports: AT-12 blank-criteria cases (student name/number/date range)')

    def set_crit(a1, v):
        if v is None or v == '':
            wb.clear_contents(REP, a1)
        elif isinstance(v, tuple):
            wb.set_value(REP, a1, '=DATE({},{},{})'.format(*v))
        else:
            wb.set_value(REP, a1, str(v))

    def try_criteria(name, num, frm, to, label):
        set_crit('B4', name)
        set_crit('B5', num)
        set_crit('B8', frm)
        set_crit('B9', to)
        wb.calculate_full_rebuild()
        line = f"  {label:<34} jobs={jobs():<4} chargeable={t('N9'):<10}"
        warn = t('A14')
        if warn != '':
            line += f' WARN: {warn}'
        print(line)

    d15, d16, d17 = (2026, 9, 15), (2026, 9, 16), (2026, 9, 17)
    try_criteria('', '', '', '', 'all blank (everything)')
    try_criteria('smith', '', '', '', 'name only, partial + lowercase')
    try_criteria('', '2203110', '', '', 'number only')
    try_criteria('', '', d16, '', 'from date only')
    try_criteria('', '', '', d16, 'to date only')
    try_criteria('', '', d16, d16, 'single-day range (incl. times)')
    try_criteria('Patel', '', d15, d17, 'name + both dates')
    try_criteria('Smith', '2203110', '', '', 'mismatched name/number (14.1)')
    try_criteria('Nobody', '', '', '', 'no matches')
    try_criteria('', '', '16/09/2026', '', 'from date as TEXT (robustness)')

    # B6 is a deliberate blank gap (locked), so the two pairs are cleared separately.
    wb.clear_contents(REP, 'B4:B5')
    wb.clear_contents(REP, 'B8:B9')

    section('Reports: new filters (Technician / Printer / Paper Stock)')
    print('  (E5/E7/E10 are dropdowns - Technician exact-matches recorded jobs, Printer/Paper stock list the whole active catalogue)')
    print('  (2026-09-27: Technician/Printer/Paper Stock shifted down one row, F4:F6 -> F5:F7, to make room for Location at F4; the inputs moved F -> E in 0.11.0)')
    wb.set_value(REP, 'E5', 'J. Okonkwo')
    wb.calculate_full_rebuild()
    print(f"  Technician = 'J. Okonkwo': jobs={jobs()}")
    check(int(jobs()) != 0, 'expected at least one match on Technician filter')
    wb.clear_contents(REP, 'E5')
    wb.set_value(REP, 'E7', 'Epson SureColor P9500')
    wb.calculate_full_rebuild()
    print(f"  Printer = 'Epson SureColor P9500': jobs={jobs()}")
    check(int(jobs()) != 0, 'expected at least one match on Printer filter')
    wb.clear_contents(REP, 'E7')

    section('Reports: filter block layout (2026-10-02)')
    expect = {
        'A4': 'Student/dept. name', 'A5': 'Student number', 'A6': 'Department', 'A8': 'From date', 'A9': 'To date',
        'A11': 'Min cost', 'A12': 'Max cost', 'A13': 'Paid',
        'D4': 'Location (print room)', 'D5': 'Technician', 'D7': 'Printer model', 'D8': 'Printer type', 'D10': 'Paper stock',
        'D11': 'Paper type', 'D12': 'Student-supplied paper',
        'A15': 'Has a problem', 'D15': 'Has notes', 'A16': 'Disregarded',
        'H7': 'Sort by', 'J7': 'Sort direction',
    }
    block = wb.texts(REP, 'A1:J16')

    def bt(a1):
        col = ord(a1[0]) - 64
        return block[int(a1[1:]) - 1][col - 1]

    for k, v in expect.items():
        check(bt(k) == v, f"{k} reads '{v}' (got '{bt(k)}')")
    for k in ('A7', 'B7', 'A10', 'B10', 'D6', 'E6', 'D9', 'E9', 'D16', 'E16'):
        check(bt(k) == '', f"{k} is a blank gap row (got '{bt(k)}')")
    check(bt('A13') == 'Paid' and bt('A12') == 'Max cost', 'Min cost, Max cost and Paid are consecutive rows (no gap above Paid, 0.10.26)')
    check(wb.column_width(REP, 5) > 10 and wb.column_width(REP, 6) < 5,
          'right-hand filter inputs are in E, beside their labels in D (0.11.0); F is the spacer')
    check('Multi-pass' in wb.validation_formula1(REP, 'E8'), 'Printer type (E8) offers Single pass / Multi-pass')
    check('Neither' in wb.validation_formula1(REP, 'B16'), 'Disregarded (B16) offers Neither (0.11.0)')

    section('Reports: Paid and Paper type filters')
    wb.calculate_full_rebuild()
    all_ = int(jobs())

    def count_with(a1, v):
        wb.set_value(REP, a1, v)
        wb.calculate_full_rebuild()
        return int(jobs())

    yes = count_with('B13', 'Yes')
    no = count_with('B13', 'No')
    wb.clear_contents(REP, 'B13')
    print(f'  all={all_} paid=Yes: {yes}  paid=No: {no}')
    check(yes + no == all_, 'Paid Yes + No should account for every job (blank counts as No)')
    roll = count_with('E11', 'Roll')
    sheet_ = count_with('E11', 'Sheet')
    wb.clear_contents(REP, 'E11')
    print(f'  paper type Roll: {roll}  Sheet: {sheet_}')
    check(roll + sheet_ == all_, 'Roll + Sheet should account for every job')
    check(roll > 0 and sheet_ > 0, 'fixture should contain both roll and sheet jobs')
    wb.calculate_full_rebuild()

    section('Reports: dropdown/date-picker UI controls')
    for a in ('E5', 'E7', 'E10', 'I7', 'K7', 'B13', 'E11'):
        vt = wb.validation_type(REP, a)
        print(f'  {a} validation type: {vt} (3 = list/dropdown)')
        check(vt == 'validate list', f'{a} should be a list dropdown')
    for a in ('B8', 'B9'):
        vt = wb.validation_type(REP, a)
        print(f'  {a} validation type: {vt} (4 = date; the native calendar picker is web-only and does not appear on desktop Excel)')
        check('date' in vt, f'{a} should be Date-type validation')
    print('  OK: Technician/Printer/Paper Stock/Paper type/Paid/Sort by/Sort direction are dropdowns; From/To date use Date validation')

    section('Reports: buttons within the first screenful')
    for b in wb.buttons(REP):
        left = int(float(b['left']))
        print(f"  {b['name']:<34} left={left}px")
        check(left <= 900, f"{b['name']} is positioned too far right ({left}px)")

    section('Reports: sort by column')
    wb.set_value(REP, 'I7', 'Qty')
    wb.set_value(REP, 'K7', 'Descending')
    wb.calculate_full_rebuild()
    qty_col = reports_col('Qty')
    qtys = [float(r[qty_col - 1] or 0) for r in spill_block(REP, 'H11')]
    print('  Qty column, sorted Descending: ' + ', '.join(f'{q:g}' for q in qtys))
    check(qtys == sorted(qtys, reverse=True), 'results were not sorted by Qty Descending')
    print('  OK: SORTBY reordered the results as requested')
    wb.clear_contents(REP, 'I7')
    wb.clear_contents(REP, 'K7')
    wb.calculate_full_rebuild()

    section('breakdowns (no criteria)')
    for a in ('AA11', 'AE11'):
        print(f'  {a}:')
        try:
            for r in spill_block(REP, a, text=True)[:6]:
                print('     ' + ' | '.join(r))
        except ExcelError:
            print('     NO SPILL: ' + t(a))

    section('Export names toggle is gone (2026-10-02); live view always shows names')
    check(t('AQ12') == '' and t('AR12') == '', 'no Export names label or dropdown at AQ12:AR12')
    wb.calculate_full_rebuild()
    name_col = reports_col('Student name')
    no_col = reports_col('Student no')
    check(name_col > 0, f'Student name column exists in the results table (col {name_col})')
    check(no_col > 0, f'Student no column exists in the results table (col {no_col})')
    sp = spill_block(REP, 'H11')
    check(any(str(r[name_col - 1] or '') != '' for r in sp), 'the live results still show Student names in the live view')
    zero_no = sum(1 for r in spill_block(REP, 'H11', text=True) if r[no_col - 1] == '0')
    check(zero_no == 0, f"a job with no student number shows blank, not '0' ({zero_no} rows show 0)")

    section('Live view shows names (again, after a rebuild)')
    wb.calculate_full_rebuild()
    sp = spill_block(REP, 'H11')
    check(any(str(r[name_col - 1] or '') != '' for r in sp), 'at least one row shows a Student name in the live view')
    print(f"  sample: '{sp[0][name_col - 1] or ''}'")
    wb.calculate_full_rebuild()

    section('Minimum-columns view: only the documented set stays visible')
    should_show = ('Date/Time', 'Location', 'Student name', 'Student no', 'Paper stock', 'Chargeable', 'Paid')
    should_hide = ('Printer', 'Qty', 'Unit', 'Area m2', 'Paper cost', 'Consumable cost', 'Gross', 'Disregarded', 'Technician', 'Notes')
    for h in should_show:
        col = reports_col(h) + 7
        check(not wb.column_hidden(REP, col), f'{h} ({col_letters(col)}) is visible')
    for h in should_hide:
        col = reports_col(h) + 7
        check(wb.column_hidden(REP, col), f'{h} ({col_letters(col)}) is hidden by default')

    section('Paid appears in the results table')
    paid_col = reports_col('Paid')
    check(paid_col > 0, f'Paid column exists in the results table (col {paid_col})')
    sample = str(spill_block(REP, 'H11')[0][paid_col - 1] or '')
    print(f"  row 1 Paid = '{sample}'")
    check(sample not in ('0', '0.0'), f"blank Paid displays as blank, not the literal '0' (got '{sample}')")

    section('3b: Go to Settings button is gone')
    caps = [b['text'] for b in wb.buttons(SUM)]
    check('Go to Settings' not in caps, "no button captioned 'Go to Settings' on Summary")
    check(any('settings sheets' in cp for cp in caps), 'the Hide/Show settings sheets toggle is still present')

    section('3a: ToggleConfigSheets stays on (or returns to) Summary')
    wb.activate(MAIN)
    a = wb.active_sheet()
    check(a == MAIN, f"active sheet is Example Print Room before toggling (got '{a}')")
    wb.run('ToggleConfigSheets')
    a = wb.active_sheet()
    check(a == 'Summary', f"active sheet is Summary after ToggleConfigSheets (got '{a}')")
    vis = wb.prop(sheet('Settings'), 'Visible')
    print(f'  Settings sheet Visible = {vis} (expect hidden, since sheets started shown)')
    check(vis != '-1', 'the four configuration sheets were actually hidden by the toggle')
    wb.run('ToggleConfigSheets')
    a = wb.active_sheet()
    check(a == 'Summary', f"still on Summary after toggling back (got '{a}')")
    check(wb.prop(sheet('Settings'), 'Visible') == '-1', 'the four configuration sheets are visible again')

    # ============================================================ Export report
    def latest_report():
        found = sorted(glob.glob(os.path.join(ctx.dir, 'PrintCosts-Report-*.xlsx')), key=os.path.getmtime)
        return found[-1] if found else None

    def export_report():
        wb.calculate_full_rebuild()
        wb.activate(REP)
        wb.run('btnExportReport')
        print(wb.run('QuietLog'))
        time.sleep(0.3)
        return latest_report()

    class Report:
        """The exported .xlsx, open in this Excel: its first sheet's used
        range as text, the header lines above the table, and the table."""

        def __init__(self, path):
            self.path = path
            self.out = Workbook(os.path.basename(path))
            open_in_excel(path)
            self.ws = osa(f'tell application "Microsoft Excel" to return name of worksheet 1 of workbook "{self.out.name}"')
            n = self.out._do('''set u to used range of worksheet (item 2 of argv) of wb
return {(first row index of u) + (count of rows of u) - 1, (first column index of u) + (count of columns of u) - 1}''',
                             self.ws, literal=True)
            self.rows, self.cols = int(n[0]), int(n[1])
            self.cells = self.out.texts(self.ws, f'A1:{col_letters(self.cols)}{self.rows}')
            self.lines = []
            for r in self.cells:
                if r[0] == 'Date/Time':
                    break
                if r[0] != '':
                    self.lines.append(r[0])
                if len(self.lines) > 30:
                    break
            self.table_row = next((i for i, r in enumerate(self.cells, 1) if r[0] == 'Date/Time'), 0)
            self.headers = []
            if self.table_row:
                for h in self.cells[self.table_row - 1]:
                    if h == '':
                        break
                    self.headers.append(h)

        def a(self, r, col=1):
            return self.cells[r - 1][col - 1] if 0 < r <= self.rows and col <= self.cols else ''

        def row_of(self, text, start=1):
            return next((r for r in range(start, self.rows + 1) if self.a(r) == text), 0)

        def page_setup(self):
            v = self.out._do('''set ps to page setup object of worksheet (item 2 of argv) of wb
return {(page orientation of ps) as text, (zoom of ps) as text, fit to pages wide of ps, print title rows of ps, right footer of ps}''',
                             self.ws, literal=True)
            d = dict(zip(('orientation', 'zoom', 'wide', 'titles', 'footer'), v))
            # AppleScript's page setup has no paper size; the file says it (OOXML's default is 1, Letter).
            with zipfile.ZipFile(self.path) as z:
                xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
            m = re.search(r'<pageSetup[^>]*\spaperSize="(\d+)"', xml)
            d['paper'] = int(m.group(1)) if m else 1
            return d

        def colour(self, what, r, col=1):
            v = self.out._do(f'return color of {what} of range (item 3 of argv) of worksheet (item 2 of argv) of wb',
                             self.ws, f'{col_letters(col)}{r}', literal=True)
            return int(v[0]) + int(v[1]) * 256 + int(v[2]) * 65536

        def close(self):
            osa(f'tell application "Microsoft Excel" to close workbook "{self.out.name}" saving no')

    section('Filtered to one printer: Printer promotes')
    wb.run('SetExportNames', False)
    wb.set_value(REP, 'E7', 'Epson SureColor P9500')
    xlsx1 = export_report()
    check(xlsx1 is not None, 'an .xlsx was produced')
    if xlsx1:
        info = Report(xlsx1)
        try:
            print('  header lines:\n    ' + '\n    '.join(info.lines))
            check(info.lines[0] == 'Print job report', 'A1 is the report title')
            for k in ('Site ID', 'Site name', 'Date range', 'Rows', 'Total chargeable', 'Still owed'):
                check(k in info.lines, f'{k} line present')
            check('Paid' in info.lines, 'Paid line present in the summary')
            check('Schema version' not in info.lines, 'Schema version NOT in the header block (moved below the table)')
            check('Generated' not in info.lines, 'Generated NOT in the header block (moved below the table)')
            printer_line = next((ln for ln in info.lines if ln.startswith('Printer:')), None)
            check(printer_line is not None, f"Printer promoted into the header ('{printer_line}')")
            if printer_line:
                check(printer_line == 'Printer: Epson SureColor P9500', f"promoted value matches the filter ('{printer_line}')")

            check(info.a(2, 1) == '' and info.a(2, 2) == '', 'row 2 is blank (gap below the title)')
            total_row = info.row_of('Total chargeable')
            check(total_row > 0, f'Total chargeable row found (row {total_row})')
            if total_row > 0:
                check(info.a(total_row - 1, 1) == '' and info.a(total_row - 1, 2) == '', 'row above Total chargeable is blank')

            table_row = info.table_row
            check(table_row > 0, f'table header row found (row {table_row})')
            print('  table columns: ' + ', '.join(info.headers))
            check('Printer' not in info.headers, 'Printer column dropped from the table')
            check('Technician' in info.headers, 'Technician column still present (was being silently dropped by the old .Hidden-bounded LastVisibleColumn)')
            check('Notes' not in info.headers, 'Notes column excluded from the printed report (2026-10-02)')

            ps = info.page_setup()
            check('landscape' in ps['orientation'], f"page is landscape (got {ps['orientation']})")
            check(ps['paper'] == 9, f"paper is A4 (got {ps['paper']})")
            check(ps['zoom'] == 'false' and int(float(ps['wide'])) == 1, 'fits one page wide')
            check(ps['titles'] == f'${table_row}:${table_row}', f"table header row repeats on every page (got '{ps['titles']}')")
            check('Page &P of &N' in str(ps['footer']), 'page numbers in the footer')
            check(info.a(total_row + 1) == 'Paid', 'Paid row follows Total chargeable')
            check(info.a(total_row + 2) == 'Still owed', 'Still owed row follows Paid')
            hdr_fill = info.colour('interior object', table_row)
            check(hdr_fill != 16777215, 'table header row is tinted')
            rr, gg, bb = hdr_fill & 255, (hdr_fill >> 8) & 255, (hdr_fill >> 16) & 255
            luma = 0.299 * rr + 0.587 * gg + 0.114 * bb
            check(luma >= 200, f'header tint stays light enough to photocopy (luma {int(luma)})')
            check('Chargeable' in info.headers, 'Chargeable column still present')
            check('Area m2' not in info.headers, 'Area m2 excluded from the export by default')

            schema_row = info.row_of('Schema version', table_row)
            check(schema_row > 0, f'Schema version found below the table (row {schema_row})')
            if schema_row > 0:
                check(info.a(schema_row - 1, 1) == '' and info.a(schema_row - 1, 2) == '', 'row above Schema version is blank')
                check(info.a(schema_row + 1) == 'Generated', 'Generated immediately follows Schema version')
                grey = 110 + 110 * 256 + 110 * 65536
                check(info.colour('font object', schema_row) == grey, 'Schema version is low-contrast grey')
                check(info.colour('font object', schema_row + 1) == grey, 'Generated is low-contrast grey')
        finally:
            info.close()

    section('Unfiltered: nothing promoted, every column present')
    # Report files are named to the minute: the first goes so the second cannot land on its path.
    if xlsx1:
        os.remove(xlsx1)
    wb.set_value(REP, 'E7', '')
    xlsx2 = export_report()
    check(xlsx2 is not None, 'a second .xlsx was produced')
    if xlsx2:
        info2 = Report(xlsx2)
        try:
            check(not any(ln.startswith('Printer:') for ln in info2.lines), 'Printer NOT promoted when unfiltered (multiple printers in the results)')
            h2 = info2.headers
            print('  table columns: ' + ', '.join(h2))
            check('Printer' in h2, 'Printer column present when unfiltered')
            check(not any(ln.startswith('Location:') for ln in info2.lines), 'Location NOT promoted when unfiltered (two rooms in the results)')
            check('Location' in h2, 'Location column present when unfiltered')
            check('Area m2' not in h2, 'Area m2 excluded from the export by default')
            check(len(h2) == 17, f'all 17 exported result columns present - Area m2 and Notes excluded (got {len(h2)})')
            name_idx = h2.index('Student name') + 1 if 'Student name' in h2 else 0
            no_idx = h2.index('Student no') + 1 if 'Student no' in h2 else 0
            check(name_idx > 0 and no_idx > 0, 'Student name/no columns still present in the export')
            name_vals = name_hyphens = no_vals = 0
            for r in range(info2.table_row + 1, info2.rows):   # stops before the Total row
                nv = info2.a(r, name_idx)
                if nv != '':
                    name_vals += 1
                    if nv == '----------':
                        name_hyphens += 1
                if info2.a(r, no_idx) != '':
                    no_vals += 1
            check(name_vals > 0 and name_vals == name_hyphens,
                  f'every exported Student name is the 10-hyphen placeholder ({name_hyphens} of {name_vals})')
            check(no_vals > 0, f'Student no is NOT redacted ({no_vals} values kept)')
        finally:
            info2.close()

    section('Export prompt answered Yes: the export includes names')
    if xlsx2:
        os.remove(xlsx2)
    wb.run('SetExportNames', True)
    h_set = wb.wb_name('SET_REPORT_HEADING')
    h_text = wb.text(*h_set)
    check(h_text == 'Print job report', f"Report heading setting exists with the default (got '{h_text}')")
    wb.call('UnlockSheet', sheet(h_set[0]))
    wb.set_value(*h_set, 'Dept of Design - Print charges')
    xlsx3 = export_report()
    check(xlsx3 is not None, 'a third .xlsx was produced')
    if xlsx3:
        info3 = Report(xlsx3)
        try:
            h3 = info3.headers
            name_idx3 = h3.index('Student name') + 1 if 'Student name' in h3 else 0
            any_name3 = name_idx3 > 0 and any(info3.a(r, name_idx3) != '' for r in range(info3.table_row + 1, info3.rows))
            promoted3 = any(ln.startswith('Student name:') for ln in info3.lines)
            check(any_name3 or promoted3, 'exported file contains student names when the prompt is answered Yes')
            check(info3.lines[0] == 'Dept of Design - Print charges', f"the heading comes from the Report heading setting (got '{info3.lines[0]}')")
        finally:
            info3.close()
    wb.run('SetExportNames', False)
    wb.run('SetQuiet', False)

print()
if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('PASS')
