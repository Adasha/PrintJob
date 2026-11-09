# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-widthcost.ps1 - same checks, same messages, same order; see
# that script for the history (0.12.2: print width scales paper cost, old rows
# keep the full-width rule through S_WidthRule).
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr, sheet  # noqa: E402

WS, TBL, SUM = 'Example Print Room', 'tblJobs_EXAMPLE', 'Summary'

with open_copy() as wb:
    wb.run('SetQuiet', True)
    dp = int(wb.wb_name_text('SET_ROUND_DP') or 2)

    def rnd(x):
        return round(x + 1e-12, dp)

    def col(h):
        return wb.table_col(WS, TBL, h)

    def cell(row, h):
        return wb.value(WS, addr(row, col(h)))

    def num(row, h):
        v = cell(row, h)
        return float(v) if v not in (None, '') else 0.0

    def text(row, h):
        return wb.cell_text(WS, row, col(h))

    header = wb.table_first_row(WS, TBL)
    first, last = header + 1, wb.table_last_row(WS, TBL)

    section('schema and column', first=True)
    check(wb.wb_name_text('SET_SCHEMA') == '1.5', f"SET_SCHEMA reads 1.5 (got '{wb.wb_name_text('SET_SCHEMA')}')")
    names = wb.table_headers(WS, TBL)
    check('S_WidthRule' in names, 'S_WidthRule column exists')
    check(names[-1] == 'S_SchemaVer' and names[-2] == 'S_WidthRule',
          'S_WidthRule sits just before S_SchemaVer, which is still last')

    section('rows from before 0.12.2 keep the full-width paper cost')
    old_narrow = 0
    roll_row = 0
    for r in range(first, last + 1):
        if text(r, 'Row Type') == 'Pass' or text(r, 'Qty') == '':
            continue
        check(text(r, 'S_WidthRule') == '', f'row {r} has no S_WidthRule')
        if text(r, 'S_Measure') == 'Sheet':
            continue
        roll_row = roll_row or r
        pw, sw = num(r, 'Print Width mm'), num(r, 'S_StockWidth_mm')
        if 0 < pw < sw:
            old_narrow += 1
            want = rnd(num(r, 'Qty') * num(r, 'S_UnitCost'))
            check(abs(num(r, 'Paper Cost') - want) < 0.005,
                  f'row {r}: narrower width ({pw} of {sw} mm) still costs the whole roll ({num(r, "Paper Cost")} = {want})')
    check(old_narrow > 0, f'the sample rows include a roll job with a narrower width ({old_narrow})')
    check(roll_row > 0, 'the sample rows include a roll job')

    prn, stk = text(roll_row, 'Printer'), text(roll_row, 'Paper Stock')
    stock_w = num(roll_row, 'S_StockWidth_mm')

    section('a new roll job pays for the width it uses')
    wb.activate(WS)
    wb.run('btnAddPrintJob')
    new = wb.table_last_row(WS, TBL)
    wb.set_value(WS, addr(new, col('Printer')), prn)
    wb.set_value(WS, addr(new, col('Paper Stock')), stk)
    wb.set_value(WS, addr(new, col('Qty')), 10)
    time.sleep(0.3)
    unit_cost = num(new, 'S_UnitCost')
    check(unit_cost > 0, f'unit cost stamped ({unit_cost})')
    check(text(new, 'S_WidthRule') == 'scaled', f"new row is marked 'scaled' (got '{text(new, 'S_WidthRule')}')")
    full = rnd(10 * unit_cost)
    check(abs(num(new, 'Paper Cost') - full) < 0.005, f'no width: full roll ({num(new, "Paper Cost")} = {full})')

    wb.set_value(WS, addr(new, col('Print Width mm')), stock_w / 2)
    time.sleep(0.3)
    half = rnd(10 * unit_cost * 0.5)
    check(abs(num(new, 'Paper Cost') - half) < 0.005, f'half the width: half the paper ({num(new, "Paper Cost")} = {half})')
    check(abs(num(new, 'Gross Cost') - (num(new, 'Paper Cost') + num(new, 'Consumable Cost') + num(new, 'Set-up Cost'))) < 0.005,
          'Gross Cost follows the scaled paper cost')

    wb.set_value(WS, addr(new, col('Print Width mm')), stock_w)
    time.sleep(0.3)
    check(abs(num(new, 'Paper Cost') - full) < 0.005, f'full width: full roll ({num(new, "Paper Cost")} = {full})')

    section('re-stamping never changes which rule a row is on')
    wb.set_value(WS, addr(new, col('Print Width mm')), stock_w / 2)
    time.sleep(0.3)
    wb.call('StampRow', sheet(WS), new - header)
    check(text(new, 'S_WidthRule') == 'scaled', 'a scaled row stays scaled after StampRow')
    check(abs(num(new, 'Paper Cost') - half) < 0.005, 'and keeps its cost')
    narrow_old = next(r for r in range(first, last + 1)
                      if text(r, 'Qty') != '' and text(r, 'S_Measure') not in ('Sheet', '') and 0 < num(r, 'Print Width mm') < num(r, 'S_StockWidth_mm'))
    before = num(narrow_old, 'Paper Cost')
    wb.call('StampRow', sheet(WS), narrow_old - header)
    wb.set_value(WS, addr(narrow_old, col('Paper Stock')), text(narrow_old, 'Paper Stock'))
    time.sleep(0.3)
    check(text(narrow_old, 'S_WidthRule') == '', 'an old row stays on the old rule after StampRow and re-picking its paper')
    check(abs(num(narrow_old, 'Paper Cost') - before) < 0.005, f'and keeps its cost ({before})')

    section('Repeat Job gives a new row the new rule')
    wb.select(WS, addr(narrow_old, col('Qty')))
    wb.run('btnRepeatJob')
    time.sleep(0.5)
    rep = wb.table_last_row(WS, TBL)
    check(text(rep, 'S_WidthRule') == 'scaled', 'a repeated old job is a new job: scaled')
    check(abs(num(rep, 'Paper Cost') - before * num(rep, 'Print Width mm') / num(rep, 'S_StockWidth_mm')) < 0.01,
          f'and costs the width it uses ({num(rep, "Paper Cost")})')

    section('the job planner estimates the same')
    wb.set_value(SUM, 'B7', prn)
    wb.set_value(SUM, 'B8', stk)
    wb.set_value(SUM, 'B9', 10)
    wb.set_value(SUM, 'B10', stock_w / 2)
    wb.call('PlannerRecalc', sheet(SUM))
    check(abs(wb.number(SUM, 'B14') - half) < 0.005, f"planner paper cost for half width ({wb.number(SUM, 'B14')} = {half})")
    wb.set_value(SUM, 'B19', wb.call('LocValue', sheet(WS), 'LOC_Name'))
    wb.run('btnPlannerAdd')
    time.sleep(0.5)
    added = wb.table_last_row(WS, TBL)
    check(text(added, 'S_WidthRule') == 'scaled', 'a planned job is scaled')
    check(abs(num(added, 'Paper Cost') - wb.number(SUM, 'B14')) < 0.005,
          f"the added row's Paper Cost equals the planner's ({num(added, 'Paper Cost')})")

    wb.run('SetQuiet', False)

finish()
