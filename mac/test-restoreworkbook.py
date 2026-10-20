# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-restoreworkbook.ps1 - same checks, same messages, same
# order; see that script for the restore-from-an-older-workbook path. Differences:
#
#   - The probe line in the old workbook's Workbook_Open, the SaveCopyAs, and
#     the "not a Print Cost workbook" file are made by a small helper module
#     put into a throwaway Builder.xlsm copy (xlmac.builder_helper):
#     editing another workbook's VBA project has to come from VBA, and this
#     keeps it out of the product.
#   - The saved copy carries macOS's quarantine flag (Excel wrote it), which
#     would make the positive-control open ask about macros, so it is cleared.
#   - "No temporary copy left behind" counts PrintCosts-restore-* in Excel's
#     own temp folder, where Restore puts it on a Mac.
#   - The three copies run one after the other in the same Excel.
#
# Drives COPIES in a temp folder, never src/PrintJob.xlsm.

import glob
import hashlib
import os
import shutil
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import (open_copy, check, section, sheet, open_in_excel, osa, ExcelError,  # noqa: E402
                   unquarantine, builder_helper, WORK_ROOT)
from fixture_annexe import add_annexe_fixture  # noqa: E402

PROBE = 'PROBE-OPEN-RAN'
EXCEL_TMP = os.path.expanduser('~/Library/Containers/com.microsoft.Excel/Data/tmp')

# No " _" continuations: modBuilder.BuilderAppend inserts line by line.
HELPER = r'''Option Explicit
Public Function ProbeAndCopy(ByVal WbName As String, ByVal Probe As String, ByVal Dest As String) As String
    Dim cm As Object, ln As Long
    On Error GoTo Fail
    Set cm = Workbooks(WbName).VBProject.VBComponents("ThisWorkbook").CodeModule
    ln = cm.ProcBodyLine("Workbook_Open", 0)
    cm.InsertLines ln + 1, "    Application.StatusBar = """ & Probe & """"
    Workbooks(WbName).SaveCopyAs Dest
    ProbeAndCopy = "ok"
    Exit Function
Fail:
    ProbeAndCopy = "ERROR " & Err.Number & ": " & Err.Description
End Function
Public Function MakeOther(ByVal Dest As String) As String
    Dim nb As Workbook
    On Error GoTo Fail
    Set nb = Workbooks.Add
    nb.SaveAs Dest, 51
    nb.Close False
    MakeOther = "ok"
    Exit Function
Fail:
    MakeOther = "ERROR " & Err.Number & ": " & Err.Description
End Function'''


def temp_copies():
    return len(glob.glob(os.path.join(EXCEL_TMP, 'PrintCosts-restore-*')))


def sum_qty(wb, ws, t):
    return round(sum(float(v) for (v,) in wb.values(ws, column_range(wb, ws, t, 'Qty')) if isinstance(v, float)), 6)


def column_range(wb, ws, t, h):
    c = xlmac.col_letters(wb.table_col(ws, t, h))
    return f'{c}{wb.table_first_row(ws, t) + 1}:{c}{wb.table_last_row(ws, t)}'


temp_before = temp_copies()
old_dir = os.path.join(WORK_ROOT, f'run-old-{os.getpid()}')
shutil.rmtree(old_dir, ignore_errors=True)
os.makedirs(old_dir)
f_old = os.path.join(old_dir, 'PrintJob.xlsm')     # same name as the fresh copy on purpose
f_other = os.path.join(old_dir, 'Other.xlsx')

section('Build the "old" workbook from copy A', first=True)
expect = {}
with open_copy() as wb:
    add_annexe_fixture(wb)
    a = 'Annexe'
    wb.call('UnlockSheet', sheet(a))
    wb.set_value(a, wb.name_addr(a, 'LOC_Dept'), 'Fine Art')
    wb.set_value(a, wb.name_addr(a, 'LOC_DefTech'), 'Test Tech')
    models = wb.table_column_texts('Printers', 'tblPrinters', 'Model')[:2]
    expect['Printers'] = ';'.join(models)
    wb.set_value(a, wb.name_addr(a, 'LOC_Printers'), expect['Printers'])
    wb.call('RelockSheet', sheet(a))

    # The real user path: events on, so the change handler rescales the roll rows.
    wb.call('UnlockSheet', sheet(a))
    wb.set_value(a, wb.name_addr(a, 'LOC_RollUnit'), 'Centimetres')
    wb.call('RelockSheet', sheet(a))
    time.sleep(0.5)
    prev = -1.0
    for _ in range(30):
        cur = sum_qty(wb, a, 'tblJobs_ANNEX')
        if cur == prev:
            break
        prev = cur
        time.sleep(0.5)

    expect.update(Dept='Fine Art', Tech='Test Tech', Unit='Centimetres',
                  Jobs=wb.table_row_count(a, 'tblJobs_ANNEX'), Qty=sum_qty(wb, a, 'tblJobs_ANNEX'),
                  Printers2=wb.table_row_count('Printers', 'tblPrinters'))
    print(f"  Annexe: {expect['Jobs']} job row(s), Qty total {expect['Qty']} (cm), printers '{expect['Printers']}'")

    # Probe: Workbook_Open of the saved copy sets the status bar.
    # The "not a Print Cost workbook" file is made here too, by Excel, so it can read it back.
    with builder_helper(HELPER) as b:
        for m, a in (('ProbeAndCopy', (wb.name, PROBE, f_old)), ('MakeOther', (f_other,))):
            out = b.run('modTestHelp.' + m, *a)
            if out != 'ok':
                raise ExcelError(f'{m}: {out}')

check(expect['Qty'] > 0, f"the old workbook has roll Qty to compare (total {expect['Qty']})")
unquarantine(f_old)
with open(f_old, 'rb') as fh:
    hash_before = hashlib.sha256(fh.read()).hexdigest()

section('Positive control: opened normally, the old workbook runs its Workbook_Open')
osa('tell application "Microsoft Excel" to set status bar to false')
open_in_excel(f_old)
try:
    time.sleep(0.5)
    sb = osa('tell application "Microsoft Excel" to return status bar')
    check(sb == PROBE, f"probe ran on a normal open (status bar '{sb}')")
finally:
    osa('tell application "Microsoft Excel" to close workbook "PrintJob.xlsm" saving no')

section('Restore fresh copy B from the old workbook')
b_ctx = open_copy(clean=False)
with b_ctx as wb:
    wb.run('SetQuiet', True)
    osa('tell application "Microsoft Excel" to set status bar to false')
    check('Annexe' not in wb.sheet_names(), 'fresh copy has no Annexe before the restore')

    wb.run('RestoreFromFileConfirmed', f_old)
    print(wb.run('QuietLog'))
    time.sleep(0.3)

    sb = osa('tell application "Microsoft Excel" to return status bar')
    check(sb != PROBE, f"old workbook's Workbook_Open did not run while it was read (status bar '{sb}')")

    names = wb.sheet_names()
    check('Annexe' in names, 'Annexe room was created from the old workbook')
    if 'Annexe' in names:
        ann = sheet('Annexe')

        def lv(n):
            return wb.call('LocValue', ann, n)

        check(lv('LOC_Code') == 'ANNEX', 'created room has code ANNEX')
        check(lv('LOC_Dept') == expect['Dept'], f"department carried over ('{lv('LOC_Dept')}')")
        check(lv('LOC_DefTech') == expect['Tech'], f"default technician carried over ('{lv('LOC_DefTech')}')")
        check(lv('LOC_Printers') == expect['Printers'], f"permitted printers carried over ('{lv('LOC_Printers')}')")
        check(lv('LOC_RollUnit') == expect['Unit'], f"roll length unit carried over ('{lv('LOC_RollUnit')}')")
        n = wb.table_row_count('Annexe', 'tblJobs_ANNEX')
        check(n == expect['Jobs'], f"Annexe holds {expect['Jobs']} job row(s) (got {n})")
        q = sum_qty(wb, 'Annexe', 'tblJobs_ANNEX')
        check(q == expect['Qty'], f"Annexe Qty total matches the old workbook, so no unit conversion crept in ({q} vs {expect['Qty']})")
        pc = wb.table_row_count('Printers', 'tblPrinters')
        check(pc == expect['Printers2'], f"tblPrinters has the old workbook's {expect['Printers2']} rows (got {pc})")

        section('Restoring again creates no second room and no duplicates')
        wb.run('RestoreFromFileConfirmed', f_old)
        time.sleep(0.3)
        count = sum(1 for s in wb.sheet_names() if s.startswith('Annexe'))
        check(count == 1, f'still one Annexe sheet (got {count})')
        n = wb.table_row_count('Annexe', 'tblJobs_ANNEX')
        check(n == expect['Jobs'], f"Annexe still {expect['Jobs']} row(s) (got {n})")

        section('A workbook that is not a Print Cost workbook is refused')
        wb.run('SetQuiet', True)
        rows_before = wb.table_row_count('Annexe', 'tblJobs_ANNEX')
        wb.run('RestoreFromFileConfirmed', f_other)
        log = wb.run('QuietLog')
        check('does not look like a Print Cost Management workbook' in log, 'refused with an explanation')
        check(wb.table_row_count('Annexe', 'tblJobs_ANNEX') == rows_before, 'nothing changed')

    wb.run('SetQuiet', False)

with open(f_old, 'rb') as fh:
    check(hashlib.sha256(fh.read()).hexdigest() == hash_before, 'the old workbook file was not modified')
temp_after = temp_copies()
check(temp_after == temp_before, f'no temporary copy left behind ({temp_after} vs {temp_before} before)')

print()
if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('PASS')
