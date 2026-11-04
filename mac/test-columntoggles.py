# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# The Show: checkboxes on a print room and the workbook-wide internal columns
# setting (0.12.0, docs/column-toggles-design.md). A test cannot click, so the
# choices go through modColumns.SetColumnGroup, which chkShowColumn also calls.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, sheet, shape, addr  # noqa: E402

MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'
PWS, PTBL = 'Printers', 'tblPrinters'
EPSON = 'Epson SureColor P9500'

GROUPS = {
    'passes': ('Passes', 'Colour'),
    'cost': ('Paper Cost', 'Consumable Cost', 'Set-up Cost', 'Gross Cost', 'Disregarded', 'Area m2'),
    'disregard': ('Disregard Paper', 'Disregard Consumable'),
    'size': ('Print Width mm', 'Sheet size'),
    'tech': ('Technician',),
    'printer': ('Printer',),
}
INTERNAL = ('Job ID', 'Row Type', 'Parent', 'Pass', 'S_SchemaVer')
ALWAYS = ('Date/Time', 'Student Name', 'Student No', 'Paper Stock', 'Unit', 'Qty', 'Chargeable Cost', 'Paid', 'Status', 'Notes')
WORKING = ('H_Issues', 'H_Ink', 'H_Setup', 'H_Notices')

with open_copy() as wb:
    wb.run('SetQuiet', True)
    wb.activate(MAIN)
    main = sheet(MAIN)

    def hid(h, ws=MAIN, tbl=TBL):
        return wb.column_hidden(ws, wb.table_col(ws, tbl, h))

    def all_hid(hs):
        return all(hid(h) for h in hs)

    def none_hid(hs):
        return not any(hid(h) for h in hs)

    def ticked(g, ws=MAIN):
        return float(wb.prop(shape(ws, f'pcb_show_{g}'), 'Value')) == 1

    def choose(g, shown):
        wb.call('SetColumnGroup', main, g, shown)

    def stored():
        return wb.text(MAIN, wb.name_addr(MAIN, 'LOC_Columns'))

    def set_printers(text):
        loc = wb.name_addr(MAIN, 'LOC_Printers')
        wb.call('UnlockSheet', main)
        wb.set_value(MAIN, loc, text)
        wb.call('RelockSheet', main)

    def set_default_tech(text):
        wb.call('UnlockSheet', main)
        wb.set_value(MAIN, wb.name_addr(MAIN, 'LOC_DefTech'), text)
        wb.call('RelockSheet', main)

    def set_colour_mode(model, mode):
        col = wb.table_col(PWS, PTBL, 'Model')
        first = wb.table_first_row(PWS, PTBL) + 1
        row = first + [m.strip() for m in wb.table_column_texts(PWS, PTBL, 'Model')].index(model)
        wb.call('UnlockSheet', sheet(PWS))
        wb.set_value(PWS, addr(row, wb.table_col(PWS, PTBL, 'Colour mode')), mode)
        wb.call('RelockSheet', sheet(PWS))
        check(col > 0, f'{model} set to {mode or "single pass"}')

    def revisit():
        """Leave and come back: activation re-applies the room."""
        wb.activate('Summary')
        time.sleep(0.2)
        wb.activate(MAIN)
        time.sleep(0.2)

    def last_job():
        return wb.table_last_row(MAIN, TBL)

    def cell(h, row):
        return wb.text(MAIN, addr(row, wb.table_col(MAIN, TBL, h)))

    section('Built state', first=True)
    check(wb.dropdown_names(MAIN) == [] or not any(n.startswith('pcb_') for n in wb.dropdown_names(MAIN)), 'the old view drop-down is gone')
    names = [b['name'] for b in wb.buttons(MAIN)]
    check(not any(n.startswith(('pcb_btnToggleCost', 'pcb_btnToggleMultiPass')) for n in names), 'the old toggle buttons are gone')
    check(wb.text(MAIN, 'D2') == '', 'row 2 no longer holds the old label')
    keys = wb.table_column_texts('Settings', 'tblSettings', 'Key')
    for k in ('LOC_REDUCED_VIEW', 'LOC_REDUCED_COLUMNS', 'LOC_MINIMAL_COLUMNS', 'COST_COLS_HIDDEN', 'MULTIPASS_COLS_HIDDEN'):
        check(k not in keys and wb.wb_name('SET_' + k) is None, f'{k} is retired')
    check(stored() == '', 'no clicks are stored on a fresh build')
    check(none_hid(ALWAYS), 'the always-visible columns show')
    check(all_hid(INTERNAL), 'the internal columns are hidden')
    check(all_hid(WORKING), 'the working columns stay hidden')
    for g in ('cost', 'disregard', 'size', 'tech'):
        check(ticked(g) and none_hid(GROUPS[g]), f'{g} defaults to shown and ticked')

    section('Each box moves exactly its own columns')
    for g, hs in GROUPS.items():
        choose(g, True)
    others = lambda g: [h for k, v in GROUPS.items() if k != g for h in v]
    for g, hs in GROUPS.items():
        choose(g, False)
        check(all_hid(hs) and not ticked(g), f'{g} off: its columns hide and the box clears')
        check(none_hid(others(g)), f'{g} off: the other groups are untouched')
        check(none_hid(ALWAYS), f'{g} off: the always-visible columns stay')
        choose(g, True)
        check(none_hid(hs) and ticked(g), f'{g} on: its columns return and the box ticks')

    section('The strip stays clear with every group off')
    for g in GROUPS:
        choose(g, False)
    check(none_hid(ALWAYS), 'every group off: the always-visible columns stay')
    start = wb.column_lefts(MAIN, 4)[3]
    x = float(wb.prop(shape(MAIN, 'pcb_lblShow'), 'Left'))
    check(abs(x - start) < 1, f'the label starts at column D ({x}, column {start})')
    prev = x + float(wb.prop(shape(MAIN, 'pcb_lblShow'), 'Width'))
    for g in GROUPS:
        left = float(wb.prop(shape(MAIN, f'pcb_show_{g}'), 'Left'))
        check(left >= prev - 0.5, f'{g} sits after the control before it ({left} >= {prev})')
        prev = left + float(wb.prop(shape(MAIN, f'pcb_show_{g}'), 'Width'))
    panel = wb.column_lefts(MAIN, 47)[46]
    check(prev <= panel, f'the strip ends before the side panel ({prev} <= {panel})')
    wb.run('btnRefreshLocations')
    time.sleep(0.5)
    check(all(all_hid(GROUPS[g]) for g in GROUPS), 'Refresh Locations keeps the choices')
    wb.run('InitialiseWorkbook')
    time.sleep(1)
    check(all(all_hid(GROUPS[g]) for g in GROUPS), 'a second setup keeps the choices')
    check(not ticked('cost'), '...and the redrawn boxes show them')
    for g in GROUPS:
        choose(g, True)

    section('Choices are per room')
    room = wb.prop(wb.call('CreatePrintRoom', 'Toggle Room', 'Facilities'), 'Name')
    wb.activate(MAIN)
    rtbl = next(t for t in wb.table_names(room) if t.startswith('tblJobs_'))
    choose('cost', False)
    check(all_hid(GROUPS['cost']), 'cost off on this room')
    check(not hid('Paper Cost', room, rtbl), 'the other room keeps its cost columns')
    room2 = wb.prop(wb.call('CreatePrintRoom', 'Toggle Room Two', 'Facilities'), 'Name')
    wb.activate(MAIN)
    rtbl2 = next(t for t in wb.table_names(room2) if t.startswith('tblJobs_'))
    check(not hid('Paper Cost', room2, rtbl2), 'a new room starts from the defaults')
    code = wb.name_addr(room2, 'LOC_Columns')
    check(wb.text(room2, code) == '', 'a new room has no stored clicks')
    choose('cost', True)

    section('The Printer rule')
    wb.clear_contents(MAIN, wb.name_addr(MAIN, 'LOC_Columns'))
    allowed = wb.text(MAIN, wb.name_addr(MAIN, 'LOC_Printers'))
    wb.call('ApplyRoomColumns', main)
    check(not hid('Printer') and ticked('printer'), 'several printers: Printer shows')
    set_printers(EPSON)
    revisit()
    check(hid('Printer') and not ticked('printer'), 'one printer: Printer hides')
    choose('printer', True)
    check(not hid('Printer'), 'a click shows it')
    revisit()
    check(not hid('Printer'), '...and the click sticks')
    set_printers(allowed)
    revisit()
    check(not hid('Printer'), 'back to several: shown')
    check('printer' not in stored(), "the rule's answer changed, so the click is dropped")
    choose('printer', False)
    check(hid('Printer'), 'several printers, click off: hidden')
    set_printers(EPSON)
    revisit()
    check(hid('Printer') and 'printer' not in stored(), 'one printer: the rule agrees, and the old click is dropped')
    set_printers(allowed)
    revisit()
    check(not hid('Printer'), 'several again: the rule shows it')

    section('The Colour passes rule')
    check(not any(m == 'multi-pass' for m in wb.table_column_texts(PWS, PTBL, 'Colour mode')), 'the built workbook has no multi-pass printer')
    check(all_hid(GROUPS['passes']) and not ticked('passes'), 'no multi-pass printer here: the pass columns hide')
    set_colour_mode(EPSON, 'multi-pass')
    revisit()
    check(none_hid(GROUPS['passes']) and ticked('passes'), 'a permitted multi-pass printer: they show')
    choose('passes', False)
    revisit()
    check(all_hid(GROUPS['passes']), 'a click hides them, and sticks')
    set_colour_mode(EPSON, '')
    revisit()
    set_colour_mode(EPSON, 'multi-pass')
    revisit()
    check(none_hid(GROUPS['passes']), "the rule's answer changed and back: the click was dropped, so they show")

    section('Showing a column when it is needed')
    set_printers(EPSON)
    revisit()
    set_default_tech('')
    choose('tech', False)
    wb.activate(MAIN)
    wb.run('btnAddPrintJob')
    time.sleep(0.3)
    job = last_job()
    check(cell('Printer', job) == EPSON, f'one permitted printer: New job fills it in (got {cell("Printer", job)!r})')
    check(hid('Printer'), '...so the Printer column stays hidden')
    check(not hid('Technician') and ticked('tech'), 'no default technician: New job brings Technician back and ticks it')
    tech = wb.table_column_texts('Print Technicians', 'tblTechnicians', 'Name')[0]
    set_default_tech(tech)
    choose('tech', False)
    wb.run('btnAddPrintJob')
    time.sleep(0.3)
    check(cell('Technician', last_job()) == tech, 'a default technician fills the new job')
    check(hid('Technician'), '...and Technician stays hidden')
    set_default_tech('')
    choose('tech', True)

    # The printer is multi-pass here, which takes sheet stock only, so a supplied roll would be refused.
    set_colour_mode(EPSON, '')
    choose('size', False)
    wb.set_value(MAIN, addr(last_job(), wb.table_col(MAIN, TBL, 'Paper Stock')), 'Supplied (Roll)')
    time.sleep(0.3)
    check(wb.text(MAIN, addr(last_job(), wb.table_col(MAIN, TBL, 'Paper Stock'))) == 'Supplied (Roll)', 'the supplied stock was accepted')
    check(none_hid(GROUPS['size']) and ticked('size'), 'student-supplied stock brings Print size back')
    set_colour_mode(EPSON, 'multi-pass')

    job = last_job()
    wb.select(MAIN, addr(job, wb.table_col(MAIN, TBL, 'Student Name')))
    choose('passes', False)
    wb.run('btnAddPass')
    time.sleep(0.3)
    check(last_job() == job + 1, 'Add pass added a row')
    check(none_hid(GROUPS['passes']) and ticked('passes'), 'Add pass brings Colour passes back')
    check(hid('Pass'), 'the Pass number stays with the internal columns')
    wb.call('ApplyRoomColumns', main)
    set_colour_mode(EPSON, '')
    revisit()
    wb.clear_contents(MAIN, wb.name_addr(MAIN, 'LOC_Columns'))
    revisit()
    check(none_hid(GROUPS['passes']), 'pass rows but no multi-pass printer: the pass columns still show')

    section('The internal columns setting')
    cell_ref = wb.wb_name('SET_SHOW_INTERNAL_COLS')
    check(cell_ref is not None and wb.text(*cell_ref) == 'No', 'the setting exists and starts at No')
    wb.call('UnlockSheet', sheet('Settings'))
    wb.set_value(*cell_ref, 'Yes')
    time.sleep(0.5)
    check(none_hid(INTERNAL), 'Yes: the internal columns show here')
    check(not hid('Job ID', room, rtbl), '...and on the other rooms')
    check(all_hid(WORKING), '...and the working columns stay hidden')
    wb.set_value(*cell_ref, 'No')
    time.sleep(0.5)
    check(all_hid(INTERNAL), 'No: they hide again')
    check(hid('Job ID', room, rtbl), '...on the other rooms too')

    wb.run('SetQuiet', False)

finish()
