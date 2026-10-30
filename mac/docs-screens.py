# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac counterpart of docs-screens.ps1 (to come): builds the demo workbook that
# wiki/images/SHOTS.md describes on a COPY of src/PrintJob.xlsm, then exports
# the Range shots as PNG.
#
#   python3 mac/docs-screens.py            demo workbook + every Range shot
#   python3 mac/docs-screens.py --layout   demo workbook, then print each
#                                          sheet's tables and shapes and export
#                                          a trial image of each sheet
#
# Images land in WORK_ROOT/docs-screens (inside Excel's container, where Excel
# asks for no file access); copy the ones you keep into wiki/images.
#
# The capture runs in a throwaway Builder.xlsm (xlmac.builder_helper), so
# nothing is added to the PrintJob copy: Range.CopyPicture (screen,
# picture), pasted into a chart SCALE times the range's size with the
# picture stretched to fill it, then Chart.Export. The picture is vector, so
# the larger chart renders sharp. Probed on Excel for Mac 16.114, 2026-10-28.
#
# Raises no dialogs: rooms are made with CreatePrintRoom (no InputBox) and the
# workbook is quiet (SetQuiet) throughout. Close any open PrintJob workbook
# first. Closes WITHOUT saving.

import argparse
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import WORK_ROOT, builder_helper, open_copy, addr, sheet, ExcelError  # noqa: E402

OUT = os.path.join(WORK_ROOT, 'docs-screens')
SCALE = 2
TECH = 'J. Rivera'

CAPTURE = '''Option Explicit
Public Function Shot(ByVal WbName As String, ByVal SheetName As String, ByVal RangeAddr As String, ByVal Path As String, ByVal Factor As Double) As String
    Dim r As Range, co As ChartObject, host As Worksheet
    On Error GoTo Fail
    Set r = Workbooks(WbName).Worksheets(SheetName).Range(RangeAddr)
    r.CopyPicture 1, -4147
    Set host = ThisWorkbook.Worksheets(1)
    Set co = host.ChartObjects.Add(0, 0, r.Width * Factor, r.Height * Factor)
    co.Activate
    co.Chart.Paste
    With co.Chart.Shapes(co.Chart.Shapes.Count)
        .LockAspectRatio = False
        .Left = 0
        .Top = 0
        .Width = co.Chart.ChartArea.Width
        .Height = co.Chart.ChartArea.Height
    End With
    co.Chart.ChartArea.Border.LineStyle = -4142
    co.Chart.Export Path
    co.Delete
    Shot = "ok"
    Exit Function
Fail:
    Shot = "ERROR " & Err.Number & ": " & Err.Description
    If Not co Is Nothing Then co.Delete
End Function
Public Function ShowNote(ByVal WbName As String, ByVal SheetName As String, ByVal RangeAddr As String) As String
    On Error GoTo Fail
    Workbooks(WbName).Worksheets(SheetName).Range(RangeAddr).Comment.Visible = True
    ShowNote = "ok"
    Exit Function
Fail:
    ShowNote = "ERROR " & Err.Number & ": " & Err.Description
End Function
Public Function Layout(ByVal WbName As String, ByVal SheetName As String) As String
    Dim ws As Worksheet, lo As ListObject, sh As Shape, s As String
    On Error GoTo Fail
    Set ws = Workbooks(WbName).Worksheets(SheetName)
    s = "used " & ws.UsedRange.Address(False, False) & " ;; "
    For Each lo In ws.ListObjects
        s = s & "table " & lo.Name & " " & lo.Range.Address(False, False) & " ;; "
    Next lo
    For Each sh In ws.Shapes
        s = s & "shape " & sh.Name & " " & sh.TopLeftCell.Address(False, False) & ":" & sh.BottomRightCell.Address(False, False)
        If Not sh.Visible Then s = s & " hidden"
        s = s & " ;; "
    Next sh
    Layout = s
    Exit Function
Fail:
    Layout = s & "ERROR " & Err.Number & ": " & Err.Description
End Function'''


def build_demo(wb):
    """The demo workbook of SHOTS.md, on top of the shipped sample."""
    wb.run('SetQuiet', True)

    for name, value in (('SET_SITE_NAME', 'University Print Bureau'), ('SET_ORG', 'Example University'),
                        ('SET_CURRENCY', '£')):
        wb.set_value(*wb.wb_name(name), value)

    # Built by is the builder's own name: the screenshots show a neutral one.
    wb.call('UnlockSheet', sheet('Settings'))
    wb.set_value(*wb.wb_name('SET_BUILT_BY'), 'Print Services')

    # The sample's technician is a real person: rename the row, and the
    # sample jobs' Technician to match.
    tws, ttbl = 'Print Technicians', 'tblTechnicians'
    real = wb.table_column_texts(tws, ttbl, 'Name')[0]
    wb.call('UnlockSheet', sheet(tws))
    wb.set_value(tws, addr(wb.table_first_row(tws, ttbl) + 1, wb.table_col(tws, ttbl, 'Name')), TECH)
    ex = 'Example Print Room'
    etbl = next(t for t in wb.table_names(ex) if t.startswith('tblJobs'))
    efirst, tcol = wb.table_first_row(ex, etbl), wb.table_col(ex, etbl, 'Technician')
    for i, v in enumerate(wb.table_column_texts(ex, etbl, 'Technician'), start=1):
        if v == real:
            wb.call('UnlockSheet', sheet(ex))
            wb.set_value(ex, addr(efirst + i, tcol), TECH)

    dws, dtbl = 'Departments', 'tblDepartments'
    wb.activate(dws)
    wb.run('AddCatalogRow', dtbl)
    row = wb.table_last_row(dws, dtbl)
    wb.call('UnlockSheet', sheet(dws))
    for h, v in (('Name', 'School of Art'), ('Aliases', 'Art School;SoA'), ('Free', 'Yes')):
        wb.set_value(dws, addr(row, wb.table_col(dws, dtbl, h)), v)

    # Colours for the Riso, then the Riso itself.
    cws, ctbl = 'Consumables', 'tblColours'
    wb.activate(cws)
    for colour, rate in (('Black', 0.30), ('Red', 0.40), ('Blue', 0.40)):
        wb.run('AddCatalogRow', ctbl)
        row = wb.table_last_row(cws, ctbl)
        wb.call('UnlockSheet', sheet(cws))
        for h, v in (('Colour', colour), ('Consumable type', 'Risograph'), ('Cost per m2', rate)):
            wb.set_value(cws, addr(row, wb.table_col(cws, ctbl, h)), v)

    pws, ptbl = 'Printers', 'tblPrinters'
    wb.activate(pws)
    wb.run('AddCatalogRow', ptbl)
    row = wb.table_last_row(pws, ptbl)
    wb.call('UnlockSheet', sheet(pws))
    for h, v in (('Model', 'Riso SF9450'), ('Consumable type', 'Risograph'), ('Cost per m2', 0.0),
                 ('Colour mode', 'multi-pass'), ('Template cost', 0.5), ('Max sheet size', 'A3'),
                 ('Active', 'Yes')):
        wb.set_value(pws, addr(row, wb.table_col(pws, ptbl, h)), v)

    tech = TECH

    ws = wb.call('CreatePrintRoom', 'Photography', '', 'PHOTO')
    room = wb.prop(ws, 'Name')
    loc = wb.name_addr(room, 'LOC_Printers')
    wb.call('UnlockSheet', sheet(room))
    wb.set_value(room, loc, 'Epson SureColor P9500;Xerox Versant 180;Riso SF9450')

    tbl = next(t for t in wb.table_names(room) if t.startswith('tblJobs'))

    def cell(row, h):
        return addr(row, wb.table_col(room, tbl, h))

    def put(row, h, v):
        wb.call('UnlockSheet', sheet(room))
        wb.set_value(room, cell(row, h), v)

    def job(fields, passes=()):
        """New job (the bottom row), its fields typed in, then its passes."""
        wb.activate(room)
        wb.call('AddPrintJob', sheet(room))
        row = wb.table_last_row(room, tbl)
        for h, v in fields:
            put(row, h, v)
        if passes:
            wb.select(room, cell(row, 'Date/Time'))
            for colour in passes:
                wb.call('AddPass', sheet(room))
                put(wb.table_last_row(room, tbl), 'Colour', colour)
        return row

    job((('Student Name', 'Alex Morgan'), ('Student No', 'S2031447'), ('Technician', tech),
         ('Printer', 'Epson SureColor P9500'), ('Paper Stock', 'Matte 180gsm 1067mm roll'),
         ('Qty', 2.0), ('Print Width mm', 900.0)))
    art = job((('Student Name', 'School of Art'), ('Technician', tech),
               ('Printer', 'Xerox Versant 180'), ('Paper Stock', 'Matte 180gsm A2 sheet'),
               ('Qty', 4.0)))
    job((('Student Name', 'Sam Patel'), ('Student No', 'S2040912'), ('Technician', tech),
         ('Printer', 'Riso SF9450'), ('Paper Stock', 'Uncoated 120gsm A3 sheet'), ('Qty', 50.0)),
        passes=('Black', 'Red'))
    put(art, 'Paid', 'Yes')

    print('demo log:', wb.run('QuietLog').strip() or '(nothing)')
    for h in ('Job ID', 'Student Name', 'Printer', 'Paper Stock', 'Row Type', 'Colour',
              'Chargeable Cost', 'Paid', 'Status'):
        print(f'  {h}:', ' | '.join(wb.table_column_texts(room, tbl, h)))
    return room


def shots(wb, b, room):
    """Every Range shot in SHOTS.md, in an order that keeps each one's state
    from leaking into the next (the Minimal view comes last)."""
    out = []

    def shot(name, sheet_name, rng, factor=SCALE):
        path = os.path.join(OUT, f'{name}.png')
        res = b.run('modTestHelp.Shot', wb.name, sheet_name, rng, path, factor)
        out.append(f'{name}: {sheet_name}!{rng} {res or "(did not run)"}')

    def table_to(sheet_name, table, last_col):
        return f'A1:{last_col}{wb.table_last_row(sheet_name, table)}'

    ex = 'Example Print Room'
    etbl = next(t for t in wb.table_names(ex) if t.startswith('tblJobs'))
    ptbl = next(t for t in wb.table_names(room) if t.startswith('tblJobs'))

    shot('Setup-Settings-1', 'Settings', 'A7:D15')
    shot('Setup-Settings-2', 'Settings', 'A1:D5')
    shot('Setup-Lookup-Lists-1', 'Settings', 'F6:N14')
    shot('Setup-Print-Technicians-1', 'Print Technicians', table_to('Print Technicians', 'tblTechnicians', 'G'))
    shot('Setup-Printers-1', 'Printers', table_to('Printers', 'tblPrinters', 'I'))
    shot('Setup-Papers-1', 'Papers', table_to('Papers', 'tblPapers', 'L'))
    shot('Setup-Consumables-1', 'Consumables', table_to('Consumables', 'tblColours', 'F'))
    shot('Setup-Departments-1', 'Departments', table_to('Departments', 'tblDepartments', 'K'))

    shot('Jobs-Recording-1', ex, f'A1:J{wb.table_first_row(ex, etbl) + 6}')
    shot('Jobs-Colour-Passes-1', room, f'A15:H{wb.table_last_row(room, ptbl)}')

    # The job with problems: its Status note shown, with the columns before it.
    last = wb.table_last_row(ex, etbl)
    status = wb.table_col(ex, etbl, 'Status')
    wb.call('UnlockSheet', sheet(ex))
    out.append('note: ' + b.run('modTestHelp.ShowNote', wb.name, ex, addr(last, status)))
    shot('Jobs-Validation-and-Status-1', ex, f'{addr(last - 4, status - 3)}:{addr(last + 6, status + 5)}')

    wb.call('LoadCatalog', True)
    wb.call('CheckSheet', sheet(ex))
    shot('Jobs-Validation-and-Status-2', ex, 'AU1:AZ9')

    shot('Jobs-Paid-Tracking-1', 'Reports', 'A1:J5')

    for a, v in (('B7', 'Epson SureColor P9500'), ('B8', 'Matte 180gsm 1067mm roll'), ('B9', 2.0),
                 ('B19', 'Photography')):
        wb.call('UnlockSheet', sheet('Summary'))
        wb.set_value('Summary', a, v)
    shot('Reports-Summary-2', 'Summary', 'A5:F21')
    shot('Reports-Summary-1', 'Summary', 'A23:O42')

    wb.call('SetViewMode', 'Minimal')
    shot('Jobs-Views-and-Units-1', room, f'A1:R{wb.table_last_row(room, ptbl)}')
    wb.call('SetViewMode', 'All')
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--layout', action='store_true')
    opts = ap.parse_args()
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT)

    with open_copy() as wb:
        room = build_demo(wb)
        with builder_helper(CAPTURE) as b:
            if opts.layout:
                for s in wb.sheet_names():
                    if s.startswith('_'):
                        continue
                    lay = b.run('modTestHelp.Layout', wb.name, s)
                    items = [x for x in lay.split(' ;; ') if x]
                    print(f'== {s}')
                    for x in items:
                        print('  ' + x)
                    if not items or not items[0].startswith('used '):
                        print('  (no layout: skipped)')
                        continue
                    used = items[0].split(' ')[1]
                    first, _, last = used.partition(':')
                    # Trial image: A1 to the used range's last column, at most 45 rows.
                    col = ''.join(c for c in (last or first) if c.isalpha())
                    path = os.path.join(OUT, f'layout-{s.replace(" ", "_")}.png')
                    print('  trial:', b.run('modTestHelp.Shot', wb.name, s, f'A1:{col}45', path, 1))
            else:
                for line in shots(wb, b, room):
                    print(line)
        print('room:', room)


try:
    main()
except ExcelError as e:
    print(f'FAIL: {e}')
    sys.exit(1)
