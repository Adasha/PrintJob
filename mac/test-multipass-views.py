# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-multipass-views.ps1 - same checks, same messages, same
# order; see that script for the reporting and view side of multi-pass
# costing. Differences:
#
#   - The toolbar panels (rectangles, not buttons) and the buttons' grey text
#     are read by a small helper module in a throwaway Builder.xlsm copy
#     (xlmac.builder_helper): AppleScript reaches neither TextFrame2 nor a
#     button's Characters.Font.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import (open_copy, check, section, addr, sheet, table, spill, builder_helper,  # noqa: E402
                   ExcelError, SUM_HDR_ROW, SUM_FIRST_ROW, SUM_TOT_ROW)

MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'
CWS, CTBL = 'Consumables', 'tblColours'
PWS, PTBL = 'Printers', 'tblPrinters'
REP, SUM = 'Reports', 'Summary'

# No " _" continuations: modBuilder.BuilderAppend inserts line by line.
HELPER = r'''Option Explicit
Public Function PanelInfo(ByVal WbName As String, ByVal SheetName As String, ByVal ShapeName As String) As String
    On Error GoTo Fail
    With Workbooks(WbName).Worksheets(SheetName).Shapes(ShapeName)
        PanelInfo = Trim$(Str$(.Left)) & "|" & Trim$(Str$(.Top)) & "|" & Trim$(Str$(.Width)) & "|" & Trim$(Str$(.Height)) & "|" & .Fill.ForeColor.RGB & "|" & .TextFrame2.TextRange.Text
    End With
    Exit Function
Fail:
    PanelInfo = "ERROR " & Err.Number & ": " & Err.Description
End Function
Public Function ButtonColour(ByVal WbName As String, ByVal SheetName As String, ByVal Prefix As String) As String
    Dim b As Object
    On Error GoTo Fail
    For Each b In Workbooks(WbName).Worksheets(SheetName).Buttons
        If Left$(b.Name, Len(Prefix)) = Prefix Then
            ButtonColour = CStr(b.Characters.Font.Color)
            Exit Function
        End If
    Next b
    ButtonColour = "ERROR button " & Prefix & " not found"
    Exit Function
Fail:
    ButtonColour = "ERROR " & Err.Number & ": " & Err.Description
End Function'''


def near(a, b):
    return abs(float(a or 0) - b) < 0.001


def helper_out(out):
    if out.startswith('ERROR') or not out:
        raise ExcelError(out or 'helper did not answer')
    return out


with builder_helper(HELPER) as helper, open_copy() as wb:
    wb.run('SetQuiet', True)
    first = wb.table_first_row(MAIN, TBL)

    def j(row_no, h):
        return addr(first + row_no, wb.table_col(MAIN, TBL, h))

    def jv(row_no, h):
        v = wb.value(MAIN, j(row_no, h))
        return '' if v is None else v

    def js(row_no, h):
        v = jv(row_no, h)
        return v if isinstance(v, str) else f'{v:g}'

    def w(row_no, h, v):
        wb.call('UnlockSheet', sheet(MAIN))
        wb.set_value(MAIN, j(row_no, h), v if isinstance(v, str) else float(v))

    def pick(row_no):
        wb.activate(MAIN)
        wb.select(MAIN, j(row_no, 'Date/Time'))

    def col_hidden(h):
        return wb.column_hidden(MAIN, wb.table_col(MAIN, TBL, h))

    def rows():
        return wb.table_row_count(MAIN, TBL)

    def spill_vals(ws, a1):
        """A spill range read once into rows of values, retried while Excel is busy."""
        for a in range(1, 5):
            try:
                ref = wb.prop(spill(ws, a1), 'Address').replace('$', '')
                return wb.values(ws, ref)
            except ExcelError:
                time.sleep(a * 2)
        raise ExcelError('spill range could not be read')

    # ---- setup -------------------------------------------------------------------
    wb.activate(CWS)
    for name, rate in (('Black', 0.30), ('Blue', 0.40)):
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

    wb.activate(MAIN)
    wb.call('AddPrintJob', sheet(MAIN))
    job = rows()
    w(job, 'Student Name', 'Views Tester')
    w(job, 'Technician', js(1, 'Technician'))
    w(job, 'Printer', 'Test RISO')
    w(job, 'Paper Stock', 'Gloss 200gsm SRA3 sheet')
    w(job, 'Qty', 100)
    for c in ('Black', 'Blue'):
        pick(job)
        wb.call('AddPass', sheet(MAIN))
        w(rows(), 'Colour', c)
    job_id = js(job, 'Job ID')
    wb.calculate_full_rebuild()
    check(near(jv(job, 'Set-up Cost'), 6.0) and near(jv(job, 'Consumable Cost'), 10.08), 'setup: job has set-up 6.00 and ink 4.32 + 5.76')

    section('Reports')
    wb.activate(REP)
    wb.run('SetQuiet', True)
    hdr = wb.texts(REP, 'H10:AA10')[0]
    print('  ' + ' | '.join(hdr))
    check(hdr.index('Passes') == hdr.index('Printer') + 1 and hdr.index('Set-up cost') == hdr.index('Consumable cost') + 1,
          'Passes follows Printer; Set-up cost follows Consumable cost')
    check(hdr.index('Paper stock') == 6 and hdr.index('Chargeable') == 15 and hdr.index('Paid') == 16,
          'Paper stock, Chargeable and Paid are where the constants say (7, 16, 17)')
    time.sleep(2)
    sv = spill_vals(REP, 'H11')
    total = int(float(wb.value(REP, 'I9')))
    check(total == sample + 1, f'Matching jobs counts jobs only ({total} of {sample + 1})')
    check(len(sv) == sample + 1, 'the results list has one row per job: no pass rows')
    pc, sc, jc = hdr.index('Passes'), hdr.index('Set-up cost'), hdr.index('Student name')
    found = None
    for r in sv:
        if str(r[jc]) == 'Views Tester':
            found = r
    check(found is not None and int(float(found[pc])) == 2 and near(found[sc], 6.0), 'the multi-pass job shows Passes 2 and Set-up cost 6.00')
    charge = float(wb.value(REP, 'N9') or 0)
    check(charge > 0, f'Matching chargeable (N9) reads a figure ({charge:g})')

    def matching():
        return int(float(wb.value(REP, 'I9')))

    wb.set_value(REP, 'E8', 'Multi-pass')
    wb.calculate_full_rebuild()
    check(matching() == 1, 'Pass type = Multi-pass finds the one multi-pass job')
    wb.set_value(REP, 'E8', 'Single pass')
    wb.calculate_full_rebuild()
    check(matching() == sample, 'Pass type = Single pass finds the rest')
    check(wb.call('HasActiveFilter', sheet(REP)) == 'True', 'a set Pass type counts as an active filter')
    wb.set_value(REP, 'E8', '')
    wb.calculate_full_rebuild()
    check(matching() == sample + 1, 'blank Pass type = everything')
    wb.set_value(REP, 'I7', 'Passes')
    wb.set_value(REP, 'K7', 'Descending')
    wb.calculate_full_rebuild()
    time.sleep(1)
    sv = spill_vals(REP, 'H11')
    check(int(float(sv[0][pc])) == 2, 'sorting by Passes (descending) puts the multi-pass job first')
    wb.set_value(REP, 'I7', '')
    wb.set_value(REP, 'K7', '')

    section('Summary')
    wb.activate(SUM)
    sh = wb.texts(SUM, f'A{SUM_HDR_ROW}:O{SUM_HDR_ROW}')[0]
    check(sh[2] == 'Passes' and sh[11] == 'Set-up cost' and sh[14] == 'Chargeable', f"Summary headers: Passes in C, Set-up cost in L ({', '.join(sh)})")
    time.sleep(1)
    sp = spill_vals(SUM, f'A{SUM_FIRST_ROW}')
    row = None
    for r in sp:
        if str(r[1]) == 'Test RISO':
            row = r
    check(row is not None and int(float(row[2])) == 2 and near(row[11], 6.0), 'the Test RISO row shows Passes 2 and Set-up cost 6.00')
    gross = sum(float(r[12] or 0) for r in sp)
    check(near(gross, float(wb.value(SUM, f'D{SUM_TOT_ROW}'))), "the table's Gross column adds up to the Gross total")

    section('The pass-column toggle is independent of the view')
    passes = ('Passes', 'Colour', 'Row Type', 'Parent', 'Pass')
    wb.activate(MAIN)

    def all_hidden(names):
        return all(col_hidden(n) for n in names)

    def none_hidden(names):
        return not any(col_hidden(n) for n in names)

    def view(mode):
        wb.call('SetViewMode', mode)

    view('All')
    check(none_hidden(passes), 'view All, toggle on: the pass columns show')
    wb.run('ToggleMultiPassColumns')
    check(all_hidden(passes), 'toggle off: they hide')
    view('All')
    check(all_hidden(passes), '...and choosing All does not switch the group back on')
    wb.run('ToggleMultiPassColumns')
    check(none_hidden(passes), 'toggle on again: they show (view All)')
    for mode in ('Reduced', 'Minimal'):
        view(mode)
        check(none_hidden(passes), f'view {mode}, toggle on: the pass columns still show (only the toggle hides them)')
    wb.run('ToggleMultiPassColumns')
    view('Reduced')
    check(all_hidden(passes), 'view Reduced, toggle off: they stay hidden')
    wb.run('ToggleMultiPassColumns')
    check(none_hidden(passes), '...and the toggle shows them again under Reduced')
    view('All')
    check(none_hidden(passes), 'back to All: shown')
    check(not col_hidden('Set-up Cost'), 'Set-up Cost is visible')
    wb.run('ToggleCostColumns')
    check(col_hidden('Set-up Cost') and col_hidden('Consumable Cost'), 'the cost toggle hides Set-up Cost with the other cost detail')
    check(none_hidden(passes), '...without touching the pass columns')
    wb.run('ToggleCostColumns')

    # A workbook restored from before 0.11.0 has the OLD Reduced-columns list (Restore overwrites it by Key).
    # The toggle must still show the columns again, and Reduced must still hide them.
    wb.call('SetSetting', 'LOC_REDUCED_COLUMNS', 'Status;Job ID;Area m2;S_SchemaVer')
    view('All')
    wb.run('ToggleMultiPassColumns')
    check(all_hidden(passes), 'old Reduced list: the toggle still hides the pass columns')
    wb.run('ToggleMultiPassColumns')
    check(none_hidden(passes), '...and the next click shows them again (it used to do nothing)')
    view('Reduced')
    view('Reduced')
    check(none_hidden(passes), '...and Reduced leaves them showing')
    view('All')

    section('Location toolbar: Print jobs | Colour passes | pass toggles | Now')

    def btn(macro):
        # Right after RepositionLocationButtons a button's name can read back as missing value; one re-read settles it.
        for _ in range(2):
            found = wb.buttons(MAIN)
            if all(b['name'] for b in found):
                break
            time.sleep(1)
        b = next((b for b in found if (b['name'] or '').startswith(f'pcb_{macro}_')), None)
        if b is None:
            raise ExcelError(f'button {macro} not found')
        return {k: (float(v) if k in ('left', 'top', 'width', 'height') else v) for k, v in b.items()}

    order = ('btnAddPrintJob', 'btnRepeatJob', 'btnRemoveRow', 'btnAddPass', 'btnRemovePass', 'btnTogglePasses', 'btnToggleAllPasses', 'btnNow')
    caps = ('New', 'Repeat', 'Delete', 'Add', 'Remove', 'Toggle passes', 'Toggle all passes', 'Now')
    for o, c in zip(order, caps):
        check(btn(o)['text'] == c, f"{o} is captioned '{c}'")
    btns = [btn(o) for o in order]
    in_order = all(btns[i]['left'] >= btns[i - 1]['left'] + btns[i - 1]['width'] - 0.01 for i in range(1, len(btns)))
    check(in_order, 'the eight toolbar buttons run left to right in that order without overlapping')
    max_btn = max(b['top'] + b['height'] for b in btns)
    hdr_top = wb.row_top_height(MAIN, first)[0]
    check(max_btn <= hdr_top, f'the toolbar stays above the table header (end {max_btn:g}, header at {hdr_top:g})')
    check(sum(1 for b in btns if b['top'] < wb.row_top_height(MAIN, 14)[0]) == 0, 'the toolbar sits below the config block')

    def panel(name):
        p = helper_out(helper.run('modTestHelp.PanelInfo', wb.name, MAIN, name)).split('|', 5)
        return {'left': float(p[0]), 'top': float(p[1]), 'width': float(p[2]), 'height': float(p[3]), 'rgb': p[4], 'text': p[5]}

    p_jobs, p_pass, p_tog = panel('pcp_Jobs'), panel('pcp_Passes'), panel('pcp_Toggles')
    check(p_jobs['text'] == 'Print jobs', "the first panel is titled 'Print jobs'")
    check(p_pass['text'] == 'Colour passes', "the second panel is titled 'Colour passes'")
    check(p_jobs['rgb'] != p_pass['rgb'], 'Print jobs and Colour passes are shaded differently')
    gap_jobs_pass = p_pass['left'] - (p_jobs['left'] + p_jobs['width'])
    gap_pass_tog = p_tog['left'] - (p_pass['left'] + p_pass['width'])
    check(0 < gap_pass_tog < gap_jobs_pass,
          f'the toggles sit closer to Colour passes ({gap_pass_tog:g}) than Colour passes to Print jobs ({gap_jobs_pass:g})')
    check(btn('btnNow')['left'] > p_tog['left'] + p_tog['width'], 'Now is to the right of the panels')
    inside = True
    for m, pp in (('btnAddPrintJob', p_jobs), ('btnRepeatJob', p_jobs), ('btnRemoveRow', p_jobs), ('btnAddPass', p_pass),
                  ('btnRemovePass', p_pass), ('btnTogglePasses', p_tog), ('btnToggleAllPasses', p_tog)):
        bb = btn(m)
        if (bb['left'] < pp['left'] or bb['left'] + bb['width'] > pp['left'] + pp['width']
                or bb['top'] < pp['top'] or bb['top'] + bb['height'] > pp['top'] + pp['height']):
            inside = False
    check(inside, 'every button sits inside its panel')

    # Rigid under a column change: widen column A, re-settle, the whole toolbar moves by the same amount.
    before = [b['left'] - btns[0]['left'] for b in btns]
    w0 = wb.column_width(MAIN, 1)
    wb.call('UnlockSheet', sheet(MAIN))
    wb.set_column_width(MAIN, 1, w0 + 12)
    wb.call('RepositionLocationButtons', sheet(MAIN))
    btns2 = [btn(o) for o in order]
    after = [b['left'] - btns2[0]['left'] for b in btns2]
    check(all(abs(x - y) <= 0.5 for x, y in zip(before, after)), "a column width change leaves the toolbar's internal spacing alone")
    wb.set_column_width(MAIN, 1, w0)
    wb.call('RepositionLocationButtons', sheet(MAIN))
    wb.call('RelockSheet', sheet(MAIN))

    section('Add pass / Remove pass grey with the selection')

    def grey_of(prefix):
        return int(helper_out(helper.run('modTestHelp.ButtonColour', wb.name, MAIN, prefix))) == 9868950

    def refresh():
        wb.call('RefreshPassButtons', sheet(MAIN))

    pick(1)
    refresh()
    check(grey_of('pcb_btnAddPass') and grey_of('pcb_btnRemovePass'), 'a single-pass job selected: Add pass and Remove pass are grey')
    check(not grey_of('pcb_btnTogglePasses'), '...Toggle passes is not (a multi-pass printer exists)')
    pick(job)
    refresh()
    check(not grey_of('pcb_btnAddPass') and not grey_of('pcb_btnRemovePass'), 'the multi-pass job selected: both are live')
    pick(job + 1)
    refresh()
    check(not grey_of('pcb_btnAddPass'), 'one of its passes selected: still live')
    check(any(b['name'].startswith('pcb_btnToggleMultiPass') for b in wb.buttons(MAIN)), 'the pass-column toggle button is drawn')

    section('No sort or filter on a location sheet')
    check(wb.prop(table(TBL), 'ShowAutoFilter') == 'False', 'the AutoFilter arrows are hidden')
    check(wb.prop(sheet(MAIN), 'Protection.AllowSorting') == 'False' and wb.prop(sheet(MAIN), 'Protection.AllowFiltering') == 'False',
          'protection withholds sorting and filtering')
    check(wb.prop(sheet(REP), 'Protection.AllowFiltering') == 'True', 'Reports keeps its filtering')

    section('The planner does not offer multi-pass printers')
    wb.run('RefreshSummaryPlanner')

    def list_items(a1):
        """The items behind a list validation: a range, or a name for one."""
        f = wb.validation_formula1(SUM, a1).lstrip('=')
        if not f:
            return []
        if '!' not in f:
            try:
                f = wb._do('return references of named item (item 2 of argv) of wb', f).lstrip('=')
            except ExcelError:
                return []
        sh_, _, ref = f.rpartition('!')
        sh_ = sh_.strip("'").replace("''", "'")
        return [c for row in wb.texts(sh_, ref.replace('$', '')) for c in row]

    items = list_items('B7')
    check('Test RISO' not in items and any('Xerox' in i for i in items),
          f"Test RISO is not in the planner's printer list (offered: {', '.join(items)})")

if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('multi-pass views checks complete')
