# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-reducedview.ps1 - same checks, same messages, same order;
# see that script for the All / Reduced / Minimal views it guards.
# Differences:
#
#   - A drop-down's items are not in Excel's AppleScript dictionary, so
#     DropDown.List(i) is read through modScript.ScriptProp on a shape token.
#   - Button.Placement comes back as 'placement free floating', COM's 3.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, shape  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'
ANNEX, ATBL = 'Annexe', 'tblJobs_ANNEX'
REDUCED = ['Status', 'Job ID', 'Area m2']
MINIMAL_EXTRA = ['Printer', 'Disregard Paper', 'Disregard Consumable', 'Print Width mm', 'Sheet size']
SIDE_PANEL = ['pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnImportLocation']

with open_copy() as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    def is_hidden(sheet, table, header):
        c = wb.table_col(sheet, table, header)
        if c < 1:
            raise RuntimeError(f"column lookup failed for '{header}'")
        return wb.column_hidden(sheet, c)

    def find_button(sheet, prefix):
        return next((b for b in wb.buttons(sheet) if b['name'].startswith(prefix)), None)

    def find_dropdown(sheet, prefix):
        return next((d for d in wb.dropdowns(sheet) if d['name'].startswith(prefix)), None)

    def column_at(sheet, left, visible_only=False):
        """First column (of 60) whose Left matches; with visible_only, the
        first VISIBLE one (a hidden column shares its Left with the next)."""
        lefts = wb.column_lefts(sheet, 60)
        hidden = wb.columns_hidden(sheet, 60) if visible_only else [False] * 60
        for c in range(1, 61):
            if abs(lefts[c - 1] - left) < 0.5 and not hidden[c - 1]:
                return c
        return -1

    def button_column(sheet, prefix, visible_only=False):
        b = find_button(sheet, prefix)
        return column_at(sheet, b['left'], visible_only) if b else 0

    def view_setting():
        return wb.wb_name_text('SET_LOC_REDUCED_VIEW')

    def view(mode):
        wb.run('SetViewMode', mode)
        time.sleep(0.3)

    # -------------------------------------------------------------- settings
    section('Settings self-provisioned', first=True)
    check(view_setting() == 'All', f"SET_LOC_REDUCED_VIEW defaults to All (got '{view_setting()}')")
    lst = wb.wb_name_text('SET_LOC_REDUCED_COLUMNS')
    extra = wb.wb_name_text('SET_LOC_MINIMAL_COLUMNS')

    def in_order(text, parts):
        pos = 0
        for p in parts:
            pos = text.find(p, pos)
            if pos < 0:
                return False
            pos += len(p)
        return True

    check(in_order(lst, ['Status', 'Job ID', 'Area m2', 'S_SchemaVer']), f"SET_LOC_REDUCED_COLUMNS holds the reduced list (got '{lst}')")
    check('Printer' not in lst, f"reduced list no longer names Printer (got '{lst}')")
    check(in_order(extra, MINIMAL_EXTRA), f"SET_LOC_MINIMAL_COLUMNS holds the extra columns (got '{extra}')")

    # ---------------------------------------------------------- initial state
    section('All: nothing in either list is hidden')
    for h in REDUCED + MINIMAL_EXTRA:
        check(not is_hidden(MAIN, TBL, h), f'{h} is visible on Example Print Room')
    h_issues_before = is_hidden(MAIN, TBL, 'H_Issues')
    check(h_issues_before, 'H_Issues (permanently hidden) starts hidden')

    section('View row: label in D2, drop-down on E, cost toggle after it; side panel toggle gone')
    d2 = wb.text(MAIN, 'D2')
    check(d2 == 'Show columns:', f"D2 reads 'Show columns:' (got '{d2}')")
    dd = find_dropdown(MAIN, 'pcb_ddViewMode')
    check(dd is not None, 'view drop-down exists')
    if dd:
        check(column_at(MAIN, dd['left']) == 5, 'drop-down sits on column E (where the All button was)')
        check(int(dd['list_count']) == 3, f"drop-down lists three modes (got {int(dd['list_count'])})")
        items = [wb.prop(shape(MAIN, dd['name']), f'List {i}') for i in (1, 2, 3)]
        check(items == ['All', 'Reduced', 'Minimal'], 'drop-down lists All, Reduced, Minimal')
        check(int(dd['list_index']) == 1, 'drop-down starts on All')
    cost = find_button(MAIN, 'pcb_btnToggleCostColumns')
    check(cost is not None, 'cost toggle exists')
    if cost and dd:
        check((cost['top'] - dd['top']) < 1 and cost['top'] < wb.row_top_height(MAIN, 3)[0], 'cost toggle is on row 2 beside the drop-down')
        check(cost['left'] >= dd['left'] + dd['width'], 'cost toggle sits to the right of the drop-down without overlapping')
    check(find_button(MAIN, 'pcb_btnViewAll') is None, 'old All button is gone')
    check(find_button(MAIN, 'pcb_btnViewReduced') is None, 'old Reduced button is gone')
    check(find_button(MAIN, 'pcb_btnViewMinimal') is None, 'old Minimal button is gone')
    check(find_button(MAIN, 'pcb_btnToggleReducedView') is None, 'old Reduce clutter button is gone')
    for p in SIDE_PANEL:
        check(button_column(MAIN, p) == 47, f'{p} sits in the side panel (column 47)')
    # Stack is packed from the top: no gap where the toggle was.
    tops = [find_button(MAIN, p)['top'] for p in SIDE_PANEL]
    step = tops[1] - tops[0]
    for i in range(1, len(tops)):
        check(abs((tops[i] - tops[i - 1]) - step) < 0.5, f'side panel step between button {i} and the previous one is even (closed up)')

    # ----------------------------------------------------------- Reduced view
    section('Reduced hides only the reduced list, on every location sheet')
    view('Reduced')
    for h in REDUCED:
        check(is_hidden(MAIN, TBL, h), f'{h} is hidden on Example Print Room')
        check(is_hidden(ANNEX, ATBL, h), f'{h} is hidden on Annexe too (workbook-wide)')
    for h in MINIMAL_EXTRA:
        check(not is_hidden(MAIN, TBL, h), f'{h} stays visible in Reduced')
    check(not is_hidden(MAIN, TBL, 'Chargeable Cost'), 'Chargeable Cost stays visible')
    check(is_hidden(MAIN, TBL, 'H_Issues') == h_issues_before, "H_Issues' state is untouched")
    check(view_setting() == 'Reduced', 'setting reads Reduced')
    check(int(find_dropdown(MAIN, 'pcb_ddViewMode')['list_index']) == 2, 'drop-down shows Reduced')
    check(int(find_dropdown(ANNEX, 'pcb_ddViewMode')['list_index']) == 2, 'drop-down shows Reduced on Annexe too')

    # ----------------------------------------------------------- Minimal view
    section('Minimal hides reduced list AND the extras')
    view('Minimal')
    for h in REDUCED + MINIMAL_EXTRA:
        check(is_hidden(MAIN, TBL, h), f'{h} is hidden in Minimal')
        check(is_hidden(ANNEX, ATBL, h), f'{h} is hidden in Minimal on Annexe')
    check(not is_hidden(MAIN, TBL, 'Chargeable Cost'), 'Chargeable Cost stays visible')
    check(is_hidden(MAIN, TBL, 'H_Issues') == h_issues_before, "H_Issues' state is untouched")
    check(view_setting() == 'Minimal', 'setting reads Minimal')

    # Printer (column E) is hidden now, so the buttons must have slid right.
    row2 = wb.texts(MAIN, 'D2:L2')[0]
    lbl_col = next((4 + i for i, t in enumerate(row2) if t == 'Show columns:'), None)
    check(lbl_col == 4, f'label still on D (Technician is visible; found col {lbl_col})')
    dd = find_dropdown(MAIN, 'pcb_ddViewMode')
    cost = find_button(MAIN, 'pcb_btnToggleCostColumns')
    c_dd = column_at(MAIN, dd['left'], visible_only=True)
    check(int(dd['list_index']) == 3, 'drop-down shows Minimal')
    check(c_dd > 0 and not wb.column_hidden(MAIN, c_dd), f'drop-down is on a visible column (col {c_dd})')
    c_cost = button_column(MAIN, 'pcb_btnToggleCostColumns', visible_only=True)
    check(c_cost > c_dd and not wb.column_hidden(MAIN, c_cost), f'cost toggle is on a later visible column (col {c_cost})')
    check(dd['left'] + dd['width'] <= cost['left'], 'drop-down and cost toggle do not overlap')
    check(abs(dd['width'] - 90) < 0.5, 'drop-down keeps its width')
    check(abs(cost['width'] - 110) < 0.5, 'cost toggle keeps its width')

    # ------------------------------------------------------------- back to All
    section('All shows everything again; drop-down back on E')
    view('All')
    for h in REDUCED + MINIMAL_EXTRA:
        check(not is_hidden(MAIN, TBL, h), f'{h} is visible again')
    check(is_hidden(MAIN, TBL, 'H_Issues') == h_issues_before, "H_Issues' state is STILL untouched")
    dd = find_dropdown(MAIN, 'pcb_ddViewMode')
    check(column_at(MAIN, dd['left']) == 5, 'drop-down back on column E')
    check(int(dd['list_index']) == 1, 'drop-down shows All again')

    # --------------------------------------------- picking from the drop-down
    section('Picking from the drop-down (ddViewMode) applies the view workbook-wide')
    wb.activate(MAIN)
    wb.set_dropdown_index(MAIN, dd['name'], 3)
    wb.run('ddViewMode')
    time.sleep(0.3)
    check(view_setting() == 'Minimal', 'ddViewMode picked Minimal')
    check(is_hidden(ANNEX, ATBL, 'Printer'), 'Minimal applied to Annexe as well')
    check(int(find_dropdown(ANNEX, 'pcb_ddViewMode')['list_index']) == 3, "Annexe's drop-down follows")
    wb.set_dropdown_index(MAIN, dd['name'], 1)
    wb.run('ddViewMode')
    time.sleep(0.3)
    check(view_setting() == 'All', 'ddViewMode picked All')
    check(not is_hidden(MAIN, TBL, 'Printer'), 'All restored the columns')

    # --------------------------------------------- settings drive both views
    section('Editing the settings changes what each view hides; Minimal still includes Reduced')
    red_cell = wb.wb_name('SET_LOC_REDUCED_COLUMNS')
    min_cell = wb.wb_name('SET_LOC_MINIMAL_COLUMNS')
    orig_list, orig_extra = lst, extra
    wb.set_value(*red_cell, 'Notes')
    wb.set_value(*min_cell, 'Paid')
    view('Reduced')
    check(is_hidden(MAIN, TBL, 'Notes'), 'Reduced hides Notes (edited list)')
    check(not is_hidden(MAIN, TBL, 'Paid'), 'Reduced does not hide Paid (extra only)')
    view('Minimal')
    check(is_hidden(MAIN, TBL, 'Notes') and is_hidden(MAIN, TBL, 'Paid'), 'Minimal hides Notes (from Reduced) and Paid (extra)')
    view('All')
    check(not is_hidden(MAIN, TBL, 'Notes') and not is_hidden(MAIN, TBL, 'Paid'), 'All shows both again')

    # Technician hidden by a test-only list: label and in-table buttons relocate.
    tech_col = wb.table_col(MAIN, TBL, 'Technician')
    check(button_column(MAIN, 'pcb_btnClearDefaults') == tech_col, 'btnClearDefaults starts anchored on Technician')
    wb.set_value(*red_cell, 'Technician')
    view('Reduced')
    check(is_hidden(MAIN, TBL, 'Technician'), 'Technician is hidden (test-only list)')
    moved = button_column(MAIN, 'pcb_btnClearDefaults', visible_only=True)
    check(moved > 0 and not wb.column_hidden(MAIN, moved), f'btnClearDefaults relocated to a visible column ({moved})')
    row2 = wb.texts(MAIN, 'D2:L2')[0]
    lbl_col = next((4 + i for i, t in enumerate(row2) if t == 'Show columns:'), None)
    check(lbl_col is not None and lbl_col > 4 and not wb.column_hidden(MAIN, lbl_col), f'label moved off hidden column D (now col {lbl_col})')
    check(wb.text(MAIN, 'D2') == '', 'D2 no longer holds the label')
    view('All')
    check(button_column(MAIN, 'pcb_btnClearDefaults') == tech_col, 'btnClearDefaults back on Technician')
    check(wb.text(MAIN, 'D2') == 'Show columns:', 'label back on D2')

    wb.set_value(*red_cell, orig_list)
    wb.set_value(*min_cell, orig_extra)

    # --------------------------- button width survives repeated view changes
    section('In-table buttons keep their width across repeated view changes')
    widths = {'pcb_btnAddPrintJob': 56, 'pcb_btnRepeatJob': 64, 'pcb_btnRemoveRow': 56, 'pcb_btnAddPass': 56,
              'pcb_btnRemovePass': 66, 'pcb_btnTogglePasses': 96, 'pcb_btnToggleAllPasses': 116, 'pcb_btnNow': 60,
              'pcb_btnClearDefaults': 110}
    for cycle in (1, 2):
        view('Minimal')
        view('All')
        for p, wdt in widths.items():
            b = find_button(MAIN, p)
            # Shape widths snap to the screen pixel grid; a point of slack, as the .ps1 allows.
            check(abs(b['width'] - wdt) < 1, f"{p} is still {wdt}pt wide after cycle {cycle} (width {b['width']})")
            check(b['placement'] == 'placement free floating', f'{p} is xlFreeFloating after cycle {cycle}')

    # ---------------------------------- side panel tracks column 47 through views
    section('Side panel stays pinned to column 47 through a Minimal/All cycle')
    col36_before = wb.column_lefts(MAIN, 47)[46]
    for p in SIDE_PANEL:
        check(find_button(MAIN, p)['left'] == col36_before, f'{p} sits on column 47 before hiding')
    view('Minimal')
    col36_hidden = wb.column_lefts(MAIN, 47)[46]
    check(col36_hidden != col36_before, 'Minimal actually moved column 47 (otherwise this proves nothing)')
    for p in SIDE_PANEL:
        check(find_button(MAIN, p)['left'] == col36_hidden, f'{p} followed column 47 while hidden')
    view('All')
    for p in SIDE_PANEL:
        check(find_button(MAIN, p)['left'] == col36_before, f'{p} back on column 47 after revealing')

    wb.run('SetQuiet', False)

finish()
