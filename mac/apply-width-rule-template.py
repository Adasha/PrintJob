# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Re-runnable template change (0.12.2): print width scales paper cost.
#
#   - Adds the job table's snapshot column S_WidthRule (schema 1.4 -> 1.5),
#     just before S_SchemaVer, which stays the last column.
#   - Rewrites the Paper Cost formula so that a roll row whose S_WidthRule is
#     "scaled" costs Qty x unit cost x (Print Width mm / stock width), capped
#     at 1. Any other value (blank on every row from before 0.12.2) keeps
#     Qty x unit cost, the whole roll width.
#
# The twelve sample rows keep S_WidthRule blank, so their figures do not move.
# Everything this adds is temporary: the first release after the late-January
# 2027 changeover removes the column and the IF around the scale factor (see
# docs/ARCHITECTURE.md, "Print width scales paper cost").
#
# PrintCosts.xlsx is binary, so the change is kept as this script rather than
# as a diff. Like apply-sample-ids-template.py it needs no Excel: the sheet,
# table and shared strings are edited inside the zip and every other part is
# copied unchanged. Works on a copy and only replaces the template once the
# copy is written. Idempotent: a second run finds nothing to change.
#
#   python3 mac/apply-width-rule-template.py [path/to/PrintCosts.xlsx]

import os
import re
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', 'src', 'PrintCosts.xlsx')
SHEET = 'xl/worksheets/sheet1.xml'      # Example Print Room
TABLE = 'xl/tables/table1.xml'          # tblJobs_EXAMPLE
SST = 'xl/sharedStrings.xml'
T = 'tblJobs_EXAMPLE'
NEW_COL = 'S_WidthRule'
MARK = 'scaled'


def ref(name):
    return f'{T}[[#This Row],[{name}]]'


OLD_PAPER = (f'IF({ref("Qty")}="","",ROUND({ref("Qty")}/IF({ref("Unit")}="cm",100,1)*{ref("S_UnitCost")},SET_ROUND_DP))')
NEW_PAPER = (f'IF({ref("Qty")}="","",ROUND({ref("Qty")}/IF({ref("Unit")}="cm",100,1)*{ref("S_UnitCost")}'
             f'*IF(AND({ref(NEW_COL)}="{MARK}",{ref("S_Measure")}<>"Sheet",N({ref("Print Width mm")})>0,N({ref("S_StockWidth_mm")})>0),'
             f'MIN(1,{ref("Print Width mm")}/{ref("S_StockWidth_mm")}),1),SET_ROUND_DP))')


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


with zipfile.ZipFile(TEMPLATE) as z:
    if any(n.lower().endswith('vbaproject.bin') for n in z.namelist()):
        sys.exit(f'{TEMPLATE} contains a VBA project; refusing to touch it')
    parts = {n: z.read(n) for n in z.namelist()}
    infos = z.infolist()

table = parts[TABLE].decode('utf-8')
if f'name="{NEW_COL}"' in table:
    print('Nothing to change: the template already has S_WidthRule.')
    sys.exit(0)

sheet = parts[SHEET].decode('utf-8')
sst = parts[SST].decode('utf-8')

# ---- shared string for the header
m = re.search(r'uniqueCount="(\d+)"', sst)
count = int(re.search(r' count="(\d+)"', sst).group(1))
unique = int(m.group(1))
sst = sst.replace('</sst>', f'<si><t>{NEW_COL}</t></si></sst>')
sst = re.sub(r' count="\d+"', f' count="{count + 1}"', sst, count=1)
sst = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{unique + 1}"', sst, count=1)
hdr_idx = unique

# ---- table: ref, column count, the new column before S_SchemaVer
if 'ref="A17:AS29"' not in table:
    sys.exit('Unexpected table range; expected A17:AS29.')
table = table.replace('ref="A17:AS29"', 'ref="A17:AT29"')
table = table.replace('<tableColumns count="45">', '<tableColumns count="46">')
ids = [int(i) for i in re.findall(r'<tableColumn id="(\d+)"', table)]
new_id = max(ids) + 1
uid = '{3D5A1C77-%04X-4B1D-9C0E-5A7E1C0D0A01}' % new_id
new_tc = f'<tableColumn id="{new_id}" xr3:uid="{uid}" name="{NEW_COL}" dataDxfId="34"/>'
anchor = re.search(r'<tableColumn id="\d+" [^>]*name="S_SchemaVer"', table)
table = table[:anchor.start()] + new_tc + table[anchor.start():]
# Paper Cost's calculated-column formula
if esc(OLD_PAPER) not in table:
    sys.exit('Paper Cost calculatedColumnFormula is not the expected text; nothing changed.')
table = table.replace(esc(OLD_PAPER), esc(NEW_PAPER))

# ---- sheet: AS -> AT, new AS, Paper Cost formulas, row spans
n_paper = sheet.count(esc(OLD_PAPER))
if n_paper == 0:
    sys.exit('No Paper Cost cell formulas found; nothing changed.')
sheet = sheet.replace(esc(OLD_PAPER), esc(NEW_PAPER))
for r in range(17, 30):
    row = re.search(r'<row r="%d"[^>]*>' % r, sheet)
    head = row.group(0).replace('spans="1:45"', 'spans="1:46"')
    sheet = sheet[:row.start()] + head + sheet[row.end():]
    cell = re.search(r'<c r="AS%d"([^>]*?)(/>|>.*?</c>)' % r, sheet)
    if not cell:
        sys.exit(f'AS{r} not found.')
    moved = f'<c r="AT{r}"{cell.group(1)}{cell.group(2)}'
    style = re.search(r's="(\d+)"', cell.group(1)).group(1)
    if r == 17:
        added = f'<c r="AS17" s="{style}" t="s"><v>{hdr_idx}</v></c>'
    else:
        added = f'<c r="AS{r}" s="{style}"/>'
    sheet = sheet[:cell.start()] + added + moved + sheet[cell.end():]

parts[SHEET] = sheet.encode('utf-8')
parts[TABLE] = table.encode('utf-8')
parts[SST] = sst.encode('utf-8')

fd, work = tempfile.mkstemp(suffix='.xlsx')
os.close(fd)
with zipfile.ZipFile(work, 'w') as out:
    for info in infos:
        out.writestr(info, parts[info.filename], compress_type=info.compress_type)
shutil.copymode(TEMPLATE, work)
shutil.move(work, TEMPLATE)
print(f'Added S_WidthRule and rewrote Paper Cost ({n_paper} cells) in {TEMPLATE}.')
