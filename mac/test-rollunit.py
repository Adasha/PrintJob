# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-rollunit.ps1 - same checks, same messages, same order; see
# that script for how LOC_RollUnit works as a display unit.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr, col_letters  # noqa: E402

MAIN = 'Example Print Room'
TBL = 'tblJobs_EXAMPLE'
DATA = '_Data'

with open_copy() as wb:
    wb.run('SetQuiet', True)

    stk_col = wb.table_col(MAIN, TBL, 'Paper Stock')
    qty_col = wb.table_col(MAIN, TBL, 'Qty')
    unit_col = wb.table_col(MAIN, TBL, 'Unit')
    area_col = wb.table_col(MAIN, TBL, 'Area m2')
    cost_col = wb.table_col(MAIN, TBL, 'Paper Cost')
    id_col = wb.table_col(MAIN, TBL, 'Job ID')

    def new_row():
        wb.activate(MAIN)
        wb.run('btnAddPrintJob')
        return wb.table_last_row(MAIN, TBL)

    def set_cell(row, col, v):
        wb.set_value(MAIN, addr(row, col), v if isinstance(v, str) else float(v))
        time.sleep(0.3)

    def v(row, col):
        return wb.value(MAIN, addr(row, col))

    def num(row, col):
        x = v(row, col)
        return 0.0 if x is None else float(x)

    def near(a, b):
        return abs(float(a or 0) - float(b)) < 0.000001

    def data_col(name):
        """_Data column (header row 9) called name, among the first 60."""
        heads = wb.texts(DATA, 'A9:BH9')[0]
        return heads.index(name) + 1 if name in heads else 0

    def data_row(job_id):
        """A job's row in _Data (data from row 10), by Job ID."""
        c = col_letters(data_col('Job ID'))
        for i, (x,) in enumerate(wb.texts(DATA, f'{c}10:{c}399')):
            if x == job_id:
                return 10 + i
        return 0

    roll_unit = wb.name_addr(MAIN, 'LOC_RollUnit')
    wb.set_value(MAIN, roll_unit, 'Metres')
    time.sleep(0.3)

    # -------------------------------------------------------- Metres baseline
    section('Metres sheet', first=True)
    r1 = new_row()
    set_cell(r1, stk_col, 'Canvas 914mm roll')
    set_cell(r1, qty_col, 2.5)
    area1 = num(r1, area_col)
    cost1 = num(r1, cost_col)
    check(v(r1, unit_col) == 'metres', 'Unit reads "metres"')
    check(near(v(r1, qty_col), 2.5), 'Qty is 2.5 as typed')
    check(cost1 > 0, f'Paper Cost calculated ({cost1})')

    r2 = new_row()
    set_cell(r2, stk_col, 'Gloss 200gsm SRA3 sheet')
    set_cell(r2, qty_col, 40)
    cost_sheet = num(r2, cost_col)

    # ------------------------------------------------- switch to Centimetres
    section('Switch to Centimetres: existing roll rows rescale')
    wb.set_value(MAIN, roll_unit, 'Centimetres')
    time.sleep(0.5)
    check(near(v(r1, qty_col), 250), f'roll Qty 2.5 m now shows 250 (got {v(r1, qty_col)})')
    check(v(r1, unit_col) == 'cm', 'roll Unit reads "cm"')
    check(near(v(r1, area_col), area1), 'Area m2 unchanged (still metres-based)')
    check(near(v(r1, cost_col), cost1), 'Paper Cost unchanged')
    check(near(v(r2, qty_col), 40), 'Sheet Qty untouched (40)')
    check(v(r2, unit_col) == 'sheets', 'Sheet Unit still "sheets"')
    check(near(v(r2, cost_col), cost_sheet), 'Sheet Paper Cost unchanged')

    # Reports' Paper type filter on a Centimetres room: see the .ps1.
    wb.calculate_full_rebuild()
    wb.set_value('Reports', 'E11', 'Roll')
    wb.calculate_full_rebuild()
    roll_jobs = int(wb.text('Reports', 'I9'))
    wb.set_value('Reports', 'E11', 'Sheet')
    wb.calculate_full_rebuild()
    sheet_jobs = int(wb.text('Reports', 'I9'))
    wb.clear_contents('Reports', 'E11')
    wb.calculate_full_rebuild()
    check(roll_jobs >= 1, f'Reports Paper type = Roll finds the roll job on a cm room (got {roll_jobs})')
    check(sheet_jobs >= 1, f'Reports Paper type = Sheet finds the sheet job on a cm room (got {sheet_jobs})')
    check(roll_jobs + sheet_jobs == int(wb.text('Reports', 'I9')), 'Roll + Sheet account for every job on a cm room')

    # ---------------------------------------------- same unit re-picked
    section('Re-picking the same unit does not rescale again')
    wb.set_value(MAIN, roll_unit, 'Centimetres')
    time.sleep(0.5)
    check(near(v(r1, qty_col), 250), 'Qty still 250')

    # ------------------------------------------- new roll row typed in cm
    section('Roll row typed in cm costs the same as its metres equivalent')
    r3 = new_row()
    set_cell(r3, stk_col, 'Canvas 914mm roll')
    set_cell(r3, qty_col, 250)
    check(near(v(r3, qty_col), 250), 'Qty stays 250 as typed (no rewrite)')
    check(v(r3, unit_col) == 'cm', 'Unit reads "cm"')
    check(near(v(r3, cost_col), cost1), 'Paper Cost equals the 2.5 m job')
    check(near(v(r3, area_col), area1), 'Area m2 equals the 2.5 m job')

    # ----------------------------------------------------------- _Data
    section('_Data (reports) always in metres')
    wb.calculate()
    q_col, u_col = data_col('Qty'), data_col('Unit')
    id1 = wb.cell_text(MAIN, r1, id_col)
    dr = data_row(id1)
    check(dr > 0, f'job {id1} found in _Data')
    if dr > 0:
        dq = wb.value(DATA, addr(dr, q_col))
        check(near(dq, 2.5), f'_Data Qty is 2.5 m (got {dq})')
        check(wb.value(DATA, addr(dr, u_col)) == 'metres', '_Data Unit is "metres"')
    id2 = wb.cell_text(MAIN, r2, id_col)
    dr2 = data_row(id2)
    if dr2 > 0:
        check(near(wb.value(DATA, addr(dr2, q_col)), 40), '_Data sheet Qty is 40')
        check(wb.value(DATA, addr(dr2, u_col)) == 'sheets', '_Data sheet Unit is "sheets"')
    else:
        check(False, f'sheet job {id2} found in _Data')

    # ----------------------------------------------------- back to Metres
    section('Back to Metres restores the lengths')
    wb.set_value(MAIN, roll_unit, 'Metres')
    time.sleep(0.5)
    check(near(v(r1, qty_col), 2.5), 'row 1 back to 2.5')
    check(near(v(r3, qty_col), 2.5), 'row 3 back to 2.5')
    check(v(r1, unit_col) == 'metres', 'Unit back to "metres"')
    check(near(v(r2, qty_col), 40), 'Sheet Qty still 40')

    wb.run('SetQuiet', False)

finish()
