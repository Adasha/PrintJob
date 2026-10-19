# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# One-time setup for the Mac tests: lets Excel read and write the folders the
# tests and the build work in.
#
#   python3 mac/grant-access.py
#
# Excel for Mac is sandboxed. Opening a workbook gives Excel that one file;
# anything the workbook's VBA then writes beside it - an export CSV, a backup,
# an Export report - goes nowhere: Save As leaves an empty file and no error,
# and the workbook's own check (rightly) reports "The export did not produce
# a file". Office's answer is GrantAccessToMultipleFiles, which shows ONE
# Grant Access prompt for a list of paths; granting a folder covers everything
# under it, and Excel remembers it. Click Grant Access when it appears.
#
# The call has to come from VBA, so it is injected into a throwaway copy of
# Builder.xlsm (never saved), as build.py injects modBuilderLive.

import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import Workbook, REPO, WORK_ROOT, BUILD_ROOT, GRANT_PATHS, open_in_excel, osa, ExcelError  # noqa: E402

FOLDERS = GRANT_PATHS   # xlmac says why both /tmp spellings
STAGE = os.path.join(BUILD_ROOT, 'grant')

CODE = '''Public Function GrantTestFolders(ByVal Paths As String) As String
    Dim ok As Boolean
    ok = GrantAccessToMultipleFiles(Split(Paths, "|"))
    GrantTestFolders = IIf(ok, "granted", "declined")
End Function'''

for f in (WORK_ROOT, BUILD_ROOT):
    os.makedirs(f, exist_ok=True)
shutil.rmtree(STAGE, ignore_errors=True)
os.makedirs(STAGE)
shutil.copy2(os.path.join(REPO, 'mac', 'Builder.xlsm'), STAGE)

b = Workbook('Builder.xlsm')
open_in_excel(os.path.join(STAGE, 'Builder.xlsm'))
try:
    for macro, args in (('modBuilder.BuilderAdd', ('Builder.xlsm', 'modGrant', 1)),
                        ('modBuilder.BuilderAppend', ('Builder.xlsm', 'modGrant', CODE))):
        out = b.run(macro, *args)
        if not out or out.startswith('ERROR'):
            sys.exit(f'{macro}: {out or "did not run"}')
    print('Excel is asking for access to:')
    for f in FOLDERS:
        print(f'  {f}')
    print('Click Grant Access in Excel (waits up to 10 minutes).', flush=True)
    result = b.run('modGrant.GrantTestFolders', '|'.join(FOLDERS))
    print(result or 'no answer')
finally:
    try:
        osa('tell application "Microsoft Excel" to close workbook "Builder.xlsm" saving no')
    except ExcelError:
        pass
    shutil.rmtree(STAGE, ignore_errors=True)
