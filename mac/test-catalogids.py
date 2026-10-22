# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-catalogids.ps1 - same checks, same messages, same order;
# see that script for the site-prefixed catalogue IDs it guards. The backup's
# CSVs are rewritten as "another site's" in place - the same files, so Excel's
# sandbox, which reads only the files it wrote, still reads them.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import glob
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, addr, col_letters, sheet, rng  # noqa: E402

SPECS = [
    {'Sheet': 'Print Technicians', 'Table': 'tblTechnicians', 'Code': 'TCH', 'Id': 'TechID', 'Name': 'Name', 'Hwm': 'TECH_ID_HWM'},
    {'Sheet': 'Printers', 'Table': 'tblPrinters', 'Code': 'PRN', 'Id': 'PrinterID', 'Name': 'Model', 'Hwm': 'PRINTER_ID_HWM'},
    {'Sheet': 'Papers', 'Table': 'tblPapers', 'Code': 'STK', 'Id': 'StockID', 'Name': 'Description', 'Hwm': 'STOCK_ID_HWM'},
]

ctx = open_copy()
with ctx as wb:
    wb.run('SetQuiet', True)
    site = wb.call('SettingText', 'SITE_ID', 'SITE')
    print(f'site code: {site}')

    def get_named(s):
        """(Id, Name) of named rows only: a blank templated row is not a record."""
        ids = wb.table_column_texts(s['Sheet'], s['Table'], s['Id'])
        names = wb.table_column_texts(s['Sheet'], s['Table'], s['Name'])
        return [(i, n) for i, n in zip(ids, names) if n.strip()]

    def num(i):
        return int(i.split('-')[-1])

    def last_id(s):
        return wb.cell_text(s['Sheet'], wb.table_last_row(s['Sheet'], s['Table']),
                            wb.table_col(s['Sheet'], s['Table'], s['Id']))

    def remove_rows_with_id(s, ids):
        ws, t = s['Sheet'], s['Table']
        wb.call('UnlockSheet', sheet(ws))
        all_ids = wb.table_column_texts(ws, t, s['Id'])
        for i in range(len(all_ids), 0, -1):
            if all_ids[i - 1] in ids:
                wb.delete_table_row(ws, t, i)
        wb.call('RelockSheet', sheet(ws))

    def hwm(s):
        return int(float(wb.call('SettingNum', s['Hwm'], 0)))

    tag = re.escape(site)

    for s in SPECS:
        section(s['Table'])
        ws, t = s['Sheet'], s['Table']
        id_c, nm_c = wb.table_col(ws, t, s['Id']), wb.table_col(ws, t, s['Name'])
        pattern = rf"^{tag}-{s['Code']}-\d{{4}}$"

        rows = get_named(s)
        check(len(rows) > 0, f'starts with named rows (got {len(rows)})')
        check(all(i.strip() for i, _ in rows), 'every named row has an ID')
        check(all(re.match(pattern, i) for i, _ in rows), f"every ID matches {site}-{s['Code']}-nnnn")
        check(len({i for i, _ in rows}) == len(rows), 'IDs are unique')
        existing = [i for i, _ in rows]

        # Add row: next ID, then delete it and add again - never reissued.
        wb.activate(ws)
        wb.run('AddCatalogRow', t)
        id1 = last_id(s)
        check(re.match(pattern, id1) is not None, f'Add row allocates a well-formed ID ({id1})')
        check(id1 not in existing, 'the new ID is not one already in use')
        c = col_letters(id_c)
        body = f'{c}{wb.table_first_row(ws, t) + 1}:{c}{wb.table_last_row(ws, t)}'
        check(wb.prop(rng(ws, body), 'Locked') == 'True', 'the whole ID column is locked, including the new row')
        check(wb.cell_format(ws, addr(wb.table_last_row(ws, t), nm_c))['locked'] is False, 'the name cell on that row is still editable')

        wb.call('UnlockSheet', sheet(ws))
        wb.delete_table_row(ws, t, wb.table_row_count(ws, t))
        wb.call('RelockSheet', sheet(ws))
        wb.run('AddCatalogRow', t)
        id2 = last_id(s)
        check(id2 != id1 and num(id2) > num(id1), f"deleting the newest row does not let its ID be reissued ({id1} deleted, next is {id2})")
        h = hwm(s)
        check(h == num(id2), f"counter {s['Hwm']} = {h} matches the last ID issued")

        # A row typed under the table: no ID until it has a name.
        wb.call('UnlockSheet', sheet(ws))
        wb.prop(xlmac.table(t), 'ListRows.Add.Index')
        wb.call('RelockSheet', sheet(ws))
        tn = wb.table_last_row(ws, t)
        wb.set_value(ws, addr(tn, nm_c), f"Typed {s['Code']} test")
        time.sleep(0.3)
        id3 = wb.cell_text(ws, tn, id_c)
        check(re.match(pattern, id3) is not None, f'a typed row gets an ID once named ({id3})')
        check(num(id3) == num(id2) + 1 if id3 else False, 'and it is the next number in sequence')

        remove_rows_with_id(s, [id2, id3])

    section("Combining another site's catalogue")
    before = {s['Table']: get_named(s) for s in SPECS}

    wb.run('BackupAll')
    csvs = glob.glob(os.path.join(ctx.dir, 'PrintCosts-*-CATALOG-*.csv'))
    check(len(csvs) == 9, f'backup wrote 9 catalogue CSVs (got {len(csvs)})')

    # Turn the backup into "another site's": same names, different site prefix.
    for s in SPECS:
        path = next(p for p in csvs if f"-CATALOG-{s['Table']}-" in os.path.basename(p))
        with open(path, 'rb') as f:
            text = f.read().decode('latin-1')
        with open(path, 'wb') as f:   # in place: Excel will read the file it wrote, not a new one
            f.write(text.replace(f"{site}-{s['Code']}-", f"OTHER-{s['Code']}-").encode('latin-1'))
    any_csv = next(p for p in csvs if '-CATALOG-tblPrinters-' in os.path.basename(p))

    # Advance one counter after the backup: a restore must not wind it back.
    tech = SPECS[0]
    wb.activate(tech['Sheet'])
    wb.run('AddCatalogRow', tech['Table'])
    bump_id = last_id(tech)
    remove_rows_with_id(tech, [bump_id])
    tech_hwm_before = hwm(tech)
    check(tech_hwm_before == num(bump_id), f'counter advanced past the backup ({bump_id})')

    wb.run('RestoreFromFileConfirmed', any_csv)
    time.sleep(0.3)

    for s in SPECS:
        after = get_named(s)
        b = before[s['Table']]
        print(f"  -- {s['Table']}: {len(b)} rows before, {len(after)} after")
        check(len(after) == 2 * len(b), f"the other site's rows were ADDED (expected {2 * len(b)})")
        check(all(o in after for o in b), 'every original row is untouched (same ID, same name)')
        check(all(any(i.startswith('OTHER-') and n == f'{o[1]} (OTHER)' for i, n in after) for o in b),
              'each clashing name was kept under a "Name (OTHER)" suffix')
        check(len({n.lower() for _, n in after}) == len(after), 'no two rows share a name')
    tech_hwm_after = hwm(tech)
    check(tech_hwm_after == tech_hwm_before, f'restore did not wind the counter back ({tech_hwm_before} -> {tech_hwm_after})')

    section('Restoring the same backup again is idempotent')
    wb.run('RestoreFromFileConfirmed', any_csv)
    time.sleep(0.3)
    for s in SPECS:
        check(len(get_named(s)) == 2 * len(before[s['Table']]), f"{s['Table']} still {2 * len(before[s['Table']])} rows")

    section('Next ID after the merge does not collide')
    wb.activate(tech['Sheet'])
    wb.run('AddCatalogRow', tech['Table'])
    id_new = last_id(tech)
    all_ids = [i for i, _ in get_named(tech)]
    check(re.match(rf'^{tag}-TCH-\d{{4}}$', id_new) is not None and id_new not in all_ids,
          f'new technician gets a fresh local-site ID ({id_new})')

    wb.run('SetQuiet', False)

print()
if xlmac._any_fail:
    print('FAIL')
    sys.exit(1)
print('PASS')
