# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Re-runnable template change (0.11.4): the sample jobs' snapshot IDs follow
# the catalogue's ID model. The twelve jobs on "Example Print Room" still
# carried the pre-0.10 IDs in S_PrinterID, S_StockID and S_TechID (PRN-003,
# STK-006, TEC-001), while the catalogue rows they were costed from have been
# UNI-PRN-0003, UNI-STK-0006 and UNI-TCH-0001 since 2026-09-29. Each old ID is
# rewritten to the new one with the same number; every pair was checked to
# name the same printer, paper or technician before this was written.
#
# PrintCosts.xlsx is binary, so the change is kept as this script rather than
# as a diff, like apply-example-code-template.ps1. Unlike that one it needs no
# Excel: the IDs are plain shared strings, so the script edits
# xl/sharedStrings.xml inside the zip and copies every other part unchanged.
# Works on a copy and only replaces the template once the copy is written.
# Idempotent: a second run finds nothing to change.
#
#   python3 mac/apply-sample-ids-template.py [path/to/PrintCosts.xlsx]

import os
import re
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', 'src', 'PrintCosts.xlsx')
SITE = 'UNI'
TABLE_CODE = {'PRN': 'PRN', 'STK': 'STK', 'TEC': 'TCH'}
OLD_ID = re.compile(r'<t>(PRN|STK|TEC)-(\d{3})</t>')


def new_id(m):
    return f'<t>{SITE}-{TABLE_CODE[m.group(1)]}-{int(m.group(2)):04d}</t>'


with zipfile.ZipFile(TEMPLATE) as z:
    if any(n.lower().endswith('vbaproject.bin') for n in z.namelist()):
        sys.exit(f'{TEMPLATE} contains a VBA project; refusing to touch it')
    sst = z.read('xl/sharedStrings.xml').decode('utf-8')
    changed, n = OLD_ID.subn(new_id, sst)
    if n == 0:
        print('Nothing to change: the sample snapshot IDs already follow the catalogue model.')
        sys.exit(0)

    fd, work = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    with zipfile.ZipFile(work, 'w') as out:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename == 'xl/sharedStrings.xml':
                data = changed.encode('utf-8')
            out.writestr(info, data, compress_type=info.compress_type)

shutil.copymode(TEMPLATE, work)
shutil.move(work, TEMPLATE)
print(f'Rewrote {n} sample snapshot IDs in {TEMPLATE}.')
