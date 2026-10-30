# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Re-runnable template change (0.11.5): the job table's Unit and Area m2
# columns are calculated, but their table-column formats (the dxf each
# tableColumn's dataDxfId points at) said locked="0". The sample rows' own
# cells are locked, so the template looked right; every row added later (New,
# Repeat, Import, Add pass, the planner) takes the column format instead, so
# on those rows both cells could be typed over, and Excel flagged them as
# unlocked cells containing formulas. Found from the green triangles in the
# 2026-10-28 Mac screenshots. This sets locked="1" in those two dxfs.
#
# Like apply-sample-ids-template.py it needs no Excel: styles.xml is edited
# inside the zip and every other part is copied unchanged. Works on a copy
# and only replaces the template once the copy is written. Idempotent: a
# second run finds nothing to change.
#
#   python3 mac/apply-calc-lock-template.py [path/to/PrintCosts.xlsx]

import os
import re
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', 'src', 'PrintCosts.xlsx')
COLUMNS = ('Unit', 'Area m2')
DXF = re.compile(r'<dxf\s*/>|<dxf>.*?</dxf>', re.S)

with zipfile.ZipFile(TEMPLATE) as z:
    if any(n.lower().endswith('vbaproject.bin') for n in z.namelist()):
        sys.exit(f'{TEMPLATE} contains a VBA project; refusing to touch it')

    ids = {}
    for name in z.namelist():
        if not name.startswith('xl/tables/'):
            continue
        x = z.read(name).decode('utf-8')
        if 'name="tblJobs_' not in x:
            continue
        for col in COLUMNS:
            m = re.search(r'<tableColumn [^>]*name="%s"[^>]*dataDxfId="(\d+)"' % re.escape(col), x)
            if m:
                ids[int(m.group(1))] = col
    if not ids:
        sys.exit('No Unit or Area m2 column format found in a tblJobs_ table')

    styles = z.read('xl/styles.xml').decode('utf-8')
    block = re.search(r'(<dxfs[^>]*>)(.*?)(</dxfs>)', styles, re.S)
    dxfs = DXF.findall(block.group(2))
    changed = []
    for i, col in sorted(ids.items()):
        new = dxfs[i].replace('<protection locked="0"', '<protection locked="1"')
        if new != dxfs[i]:
            dxfs[i] = new
            changed.append(f'{col} (dxf {i})')
    if not changed:
        print('Nothing to change: Unit and Area m2 are already locked in their column formats.')
        sys.exit(0)

    # Rebuild the dxfs block from the edited list; nothing but dxf elements
    # sits inside it.
    styles = styles[:block.start(2)] + ''.join(dxfs) + styles[block.end(2):]

    fd, work = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    with zipfile.ZipFile(work, 'w') as out:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename == 'xl/styles.xml':
                data = styles.encode('utf-8')
            out.writestr(info, data, compress_type=info.compress_type)

shutil.copymode(TEMPLATE, work)
shutil.move(work, TEMPLATE)
print(f'Locked {", ".join(changed)} in {TEMPLATE}.')
