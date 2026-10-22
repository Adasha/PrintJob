# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-importunits.ps1 - same checks, same messages, same order;
# see that script for the mixed-unit import it guards. The CSVs are imported
# exactly where and as Excel wrote them: Excel's sandbox will not read a copy.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import glob
import os
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

    def near(a, b):
        return abs(float(a or 0) - float(b)) < 0.000001

    def set_cell(ws, row, c, v):
        wb.set_value(ws, addr(row, c), v if isinstance(v, str) else float(v))
        time.sleep(0.3)

    def v(ws, row, c):
        return wb.value(ws, addr(row, c))

    def num(ws, row, c):
        x = v(ws, row, c)
        return 0.0 if x is None else float(x)

    def cc(ws, table, name):
        return wb.table_col(ws, table, name)

    def new_row(ws, table):
        wb.activate(ws)
        wb.run('btnAddPrintJob')
        return wb.table_last_row(ws, table)

    def find_row(ws, table, job_id):
        all_ids = wb.table_column_texts(ws, table, 'Job ID')
        return wb.table_first_row(ws, table) + 1 + all_ids.index(job_id) if job_id in all_ids else 0

    def newest_csv(code):
        found = sorted(glob.glob(os.path.join(ctx.dir, f'PrintCosts-*-{code}-*.csv')), key=os.path.getmtime)
        return found[-1] if found else None

    def import_into(ws, path):
        rows = wb.call('ReadImportRows', path)
        wb.call('ApplyImportConfirmed', sheet(ws), rows)
        wb.calculate_full_rebuild()
        time.sleep(0.3)

    wb.set_value(MAIN, wb.name_addr(MAIN, 'LOC_RollUnit'), 'Metres')
    wb.set_value(ANNEX, wb.name_addr(ANNEX, 'LOC_RollUnit'), 'Centimetres')
    time.sleep(0.5)

    # ---------------------------------------- jobs typed in Main (metres)
    section('Setup: roll 2.5 m and sheet x40 in the Metres room', first=True)
    m_stk, m_qty, m_unit = cc(MAIN, MTBL, 'Paper Stock'), cc(MAIN, MTBL, 'Qty'), cc(MAIN, MTBL, 'Unit')
    m_area, m_cost, m_id = cc(MAIN, MTBL, 'Area m2'), cc(MAIN, MTBL, 'Paper Cost'), cc(MAIN, MTBL, 'Job ID')

    r_roll = new_row(MAIN, MTBL)
    set_cell(MAIN, r_roll, m_stk, 'Canvas 914mm roll')
    set_cell(MAIN, r_roll, m_qty, 2.5)
    roll_id = wb.cell_text(MAIN, r_roll, m_id)
    roll_area, roll_cost = num(MAIN, r_roll, m_area), num(MAIN, r_roll, m_cost)
    check(roll_cost > 0, f'roll Paper Cost calculated ({roll_cost})')

    r_sheet = new_row(MAIN, MTBL)
    set_cell(MAIN, r_sheet, m_stk, 'Gloss 200gsm SRA3 sheet')
    set_cell(MAIN, r_sheet, m_qty, 40)
    sheet_id = wb.cell_text(MAIN, r_sheet, m_id)
    sheet_cost = num(MAIN, r_sheet, m_cost)

    # --------------------------------------------- metres file -> cm room
    section('Metres file imported into the Centimetres room')
    wb.activate(MAIN)
    wb.run('btnExport')
    csv_m = newest_csv('EXAMPLE')
    check(csv_m is not None, 'Main export written')
    import_into(ANNEX, csv_m)

    a_qty, a_unit, a_area = cc(ANNEX, ATBL, 'Qty'), cc(ANNEX, ATBL, 'Unit'), cc(ANNEX, ATBL, 'Area m2')
    a_cost, a_id, a_stk = cc(ANNEX, ATBL, 'Paper Cost'), cc(ANNEX, ATBL, 'Job ID'), cc(ANNEX, ATBL, 'Paper Stock')

    ar = find_row(ANNEX, ATBL, roll_id)
    check(ar > 0, f'roll job {roll_id} landed in Annexe')
    if ar > 0:
        check(near(v(ANNEX, ar, a_qty), 250), f'roll Qty 2.5 m became 250 cm (got {v(ANNEX, ar, a_qty)})')
        check(v(ANNEX, ar, a_unit) == 'cm', 'Unit reads "cm"')
        check(near(v(ANNEX, ar, a_area), roll_area), 'Area m2 unchanged')
        check(near(v(ANNEX, ar, a_cost), roll_cost), 'Paper Cost unchanged')
    as_ = find_row(ANNEX, ATBL, sheet_id)
    check(as_ > 0, f'sheet job {sheet_id} landed in Annexe')
    if as_ > 0:
        check(near(v(ANNEX, as_, a_qty), 40), 'sheet Qty not converted (40)')
        check(v(ANNEX, as_, a_unit) == 'sheets', 'sheet Unit "sheets"')
        check(near(v(ANNEX, as_, a_cost), sheet_cost), 'sheet Paper Cost unchanged')

    # --------------------------------------------- cm file -> metres room
    section('Centimetres file imported into the Metres room')
    r_cm = new_row(ANNEX, ATBL)
    set_cell(ANNEX, r_cm, a_stk, 'Canvas 914mm roll')
    set_cell(ANNEX, r_cm, a_qty, 150)
    cm_id = wb.cell_text(ANNEX, r_cm, a_id)
    cm_area, cm_cost = num(ANNEX, r_cm, a_area), num(ANNEX, r_cm, a_cost)
    check(v(ANNEX, r_cm, a_unit) == 'cm', 'Annexe row typed in cm reads "cm"')
    check(cm_cost > 0, f'cm job Paper Cost calculated ({cm_cost})')
    wb.activate(ANNEX)
    wb.run('btnExport')
    csv_a = newest_csv('ANNEX')
    check(csv_a is not None, 'Annexe export written')
    import_into(MAIN, csv_a)

    mr = find_row(MAIN, MTBL, cm_id)
    check(mr > 0, f'cm job {cm_id} landed in Main')
    if mr > 0:
        check(near(v(MAIN, mr, m_qty), 1.5), f'roll Qty 150 cm became 1.5 m (got {v(MAIN, mr, m_qty)})')
        check(v(MAIN, mr, m_unit) == 'metres', 'Unit reads "metres"')
        check(near(v(MAIN, mr, m_area), cm_area), 'Area m2 unchanged')
        check(near(v(MAIN, mr, m_cost), cm_cost), 'Paper Cost unchanged')
    # The cm file also carried Main's original jobs (held in Annexe at 250 cm);
    # overwriting them must land back at the original metres, not 250 or 0.025.
    rr = find_row(MAIN, MTBL, roll_id)
    check(rr > 0 and near(v(MAIN, rr, m_qty), 2.5), 'original 2.5 m roll round-trips through cm back to 2.5')
    if rr > 0:
        check(near(v(MAIN, rr, m_cost), roll_cost), 'round-tripped Paper Cost unchanged')
    rs = find_row(MAIN, MTBL, sheet_id)
    check(rs > 0 and near(v(MAIN, rs, m_qty), 40), 'sheet job still 40 after the round trip')

    # ----------------------------------------------- same-unit control
    section('Same unit: Metres file back into the Metres room')
    import_into(MAIN, csv_m)
    rr = find_row(MAIN, MTBL, roll_id)
    check(rr > 0 and near(v(MAIN, rr, m_qty), 2.5), 'metres into metres: Qty untouched (2.5)')
    if rr > 0:
        check(v(MAIN, rr, m_unit) == 'metres', 'Unit still "metres"')

    wb.run('SetQuiet', False)

finish()
