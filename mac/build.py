# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac counterpart of build.ps1: builds src/PrintJob.xlsm from
# src/PrintCosts.xlsx + src/PrintCosts-VBA/*, the same steps in the same order.
# See build.ps1 for why each step is there; only the Mac differences are
# explained here.
#
#   python3 mac/build.py
#
# Needs mac/Builder.xlsm (made by hand once - SETUP.md, "Building on a Mac")
# and Excel > Preferences > Security > "Trust access to the VBA project
# object model".
#
# The VBA project is only reachable from VBA inside Excel, so the module
# clearing and pasting goes through mac/Builder.xlsm's modBuilder; the code
# travels as text through AppleScript, not by VBComponents.Import, because
# Excel's sandbox would stop VBA reading the OneDrive folder. Everything
# else is AppleScript, as the rest of xlmac.py is.
#
# A compile error in the imported code shows a modal dialog in Excel the
# moment the first macro runs, and the AppleEvent waiting on it times out. If
# InitialiseWorkbook times out, look at Excel.

import datetime
import glob
import os
import re
import shutil
import subprocess
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import osa, ExcelError, Workbook, REPO, BUILD_ROOT, open_workbook_names, open_in_excel  # noqa: E402

ROOT = os.path.join(REPO, 'src')
VBA = os.path.join(ROOT, 'PrintCosts-VBA')
SRC = os.path.join(ROOT, 'PrintCosts.xlsx')
DST = os.path.join(ROOT, 'PrintJob.xlsm')
BUILDER = os.path.join(REPO, 'mac', 'Builder.xlsm')
BUILDER_BAS = os.path.join(REPO, 'mac', 'modBuilder.bas')

CHUNK_LINES = 300   # lines per BuilderAppend call; keeps each AppleEvent small


def die(msg):
    print(msg, file=sys.stderr)
    sys.exit(1)


def open_wb(path):
    open_in_excel(path, timeout=120)   # xlmac says why not AppleScript's open


def close_wb(name):
    try:
        osa('''on run argv
tell application "Microsoft Excel" to close workbook (item 1 of argv) saving no
end run''', name)
    except ExcelError:
        pass


LIVE = 'modBuilderLive'


def builder_call(builder, macro, *args, module=LIVE):
    """A Builder function, module-qualified: modBuilderLive (this checkout's
    modBuilder.bas) once it is loaded, the copy's own modBuilder before.
    Empty means Excel did not find it; ERROR is its own failure."""
    out = builder.run(f'{module}.{macro}', *args)
    if not out:
        die(f'{module}.{macro} did not run - is Builder.xlsm open, and its modBuilder intact?')
    if out.startswith('ERROR'):
        die(f'{macro}: {out}')
    return out


def load_live_builder(builder):
    """modBuilder.bas, loaded into the throwaway Builder.xlsm copy as
    modBuilderLive, by the copy's own (possibly older) modBuilder - so
    Builder.xlsm never needs remaking by hand. See modBuilder.bas."""
    _, _, lines = module_source(BUILDER_BAS)
    if any(ln.rstrip().endswith(' _') for ln in lines):
        die('modBuilder.bas has a " _" continuation line, which BuilderAppend cannot load. See its header.')
    builder_call(builder, 'BuilderAdd', builder.name, LIVE, 1, module='modBuilder')
    builder_call(builder, 'BuilderAppend', builder.name, LIVE, '\n'.join(lines), module='modBuilder')
    got = builder_call(builder, 'BuilderLineCount', builder.name, LIVE, module='modBuilder')
    if int(got) != len(lines):
        die(f'{LIVE}: sent {len(lines)} lines, module holds {got}')


def module_source(path):
    """(component name, kind, code lines) for a .bas/.cls - the Attribute and
    VERSION/BEGIN header Excel writes on export is dropped, as Import drops it."""
    with open(path, encoding='utf-8') as f:
        lines = f.read().splitlines()
    name = None
    for ln in lines:
        m = re.match(r'Attribute VB_Name = "([^"]+)"', ln)
        if m:
            name = m.group(1)
            break
    if not name:
        die(f'No VB_Name in {path}')
    start = next((i for i, ln in enumerate(lines) if not re.match(r'(VERSION |BEGIN|END$|  |Attribute )', ln)), None)
    kind = 2 if path.endswith('.cls') else 1
    return name, kind, lines[start:]


def send(builder, wb_name, mod, lines):
    """A whole module, in pieces into Builder's buffer, then one commit
    (modBuilder.bas, BuilderBuffer, says why it has to be one)."""
    for i in range(0, len(lines), CHUNK_LINES):
        builder_call(builder, 'BuilderBuffer', '\n'.join(lines[i:i + CHUNK_LINES]))
    got = builder_call(builder, 'BuilderCommit', wb_name, mod).split()[1]
    if int(got) != len(lines):
        die(f'{mod}: sent {len(lines)} lines, module holds {got}')


def git_suffix():
    try:
        sha = subprocess.run(['git', '-C', REPO, 'rev-parse', '--short', 'HEAD'],
                             capture_output=True, text=True, check=True).stdout.strip()
        dirty = subprocess.run(['git', '-C', REPO, 'status', '--porcelain'],
                               capture_output=True, text=True).stdout.strip()
        return '-g' + sha + ('+' if dirty else '')
    except (OSError, subprocess.CalledProcessError):
        return ''


def keep_newest(pattern, n):
    for p in sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)[n:]:
        os.remove(p)


# --------------------------------------------------------------- checks
for p in (SRC, BUILDER, BUILDER_BAS):
    if not os.path.exists(p):
        die(f'Not found: {p}' + (' - see SETUP.md, "Building on a Mac"' if p == BUILDER else ''))

# Excel for Mac takes no write lock a script can test for, so the check is
# by name: a build that copies over an open PrintJob.xlsm leaves Excel
# holding a file that is not there any more.
clash = [n for n in open_workbook_names() if n.lower().startswith(('printjob', 'printcosts', 'builder'))]
if clash:
    die(f'Close these in Excel and run this again: {", ".join(clash)}')

with zipfile.ZipFile(SRC) as z:
    if 'xl/vbaProject.bin' in z.namelist():
        die(f'{os.path.basename(SRC)} contains a VBA project, so it is not a clean source. '
            'Restore it from OneDrive version history.')

with open(BUILDER_BAS, encoding='utf-8') as f:
    want_builder = 'builder ' + re.search(r'BUILDER_VERSION As String = "([^"]+)"', f.read()).group(1)

stamp = datetime.datetime.now()
if os.path.exists(DST):
    bak = os.path.join(ROOT, 'PrintJob.' + stamp.strftime('%Y%m%d-%H%M%S') + '.bak.xlsm')
    shutil.copy2(DST, bak)
    print(f'backed up existing xlsm -> {os.path.basename(bak)}')

# A build must never mutate its own input: everything Excel opens is a copy
# outside OneDrive (build.ps1 has the history).
os.makedirs(BUILD_ROOT, exist_ok=True)
# One fixed folder, reused by every build, for the same reason as
# xlmac.open_copy's: the sandbox's file-access grant is per folder.
work_dir = os.path.join(BUILD_ROOT, 'work')
shutil.rmtree(work_dir, ignore_errors=True)
os.makedirs(work_dir)
work = os.path.join(work_dir, 'PrintCostsBuild.xlsx')
work_dst = os.path.join(work_dir, 'PrintCostsBuild.xlsm')
shutil.copy2(SRC, work)
shutil.copy2(BUILDER, os.path.join(work_dir, 'Builder.xlsm'))
src_stamp_before = os.path.getmtime(SRC)

wb_name = os.path.basename(work)
builder = Workbook('Builder.xlsm')
built = Workbook(wb_name)
opened = []
try:
    open_wb(os.path.join(work_dir, 'Builder.xlsm'))
    opened.append('Builder.xlsm')
    if not builder.run('modBuilder.BuilderReady'):
        die('Builder.xlsm has no working modBuilder. Remake it (SETUP.md, "Building on a Mac"), '
            'and check "Trust access to the VBA project object model" is ticked.')
    load_live_builder(builder)
    got = builder.run(f'{LIVE}.BuilderReady')
    if got != want_builder:
        die(f"loaded modBuilder.bas answers '{got}', expected '{want_builder}'")

    open_wb(work)
    opened.append(wb_name)

    # --- clear any previously imported components ---------------------------
    print(f'cleared  {builder_call(builder, "BuilderClear", wb_name)}')

    # --- import ------------------------------------------------------------
    files = sorted(p for p in glob.glob(os.path.join(VBA, '*'))
                   if p.endswith(('.bas', '.cls')) and os.path.basename(p) != 'ThisWorkbook.cls')
    if not files:
        die(f'No .bas/.cls files found in {VBA}')
    for p in files:
        name, kind, lines = module_source(p)
        builder_call(builder, 'BuilderAdd', wb_name, name, kind)
        send(builder, wb_name, name, lines)
        print(f'imported {os.path.basename(p)}')

    # --- ThisWorkbook ------------------------------------------------------
    tw = os.path.join(VBA, 'ThisWorkbook.cls')
    with open(tw, encoding='utf-8') as f:
        tw_lines = f.read().splitlines()
    start = next((i for i, ln in enumerate(tw_lines) if re.match(r'\s*Option Explicit', ln)), None)
    if start is None:
        die("No 'Option Explicit' line in ThisWorkbook.cls")
    send(builder, wb_name, 'ThisWorkbook', tw_lines[start:])
    print(f'pasted   ThisWorkbook.cls ({len(tw_lines) - start} lines)')

    # Builder.xlsm is done. Every call is workbook-qualified anyway
    # (Workbook.run), but setup should run with nothing else open, as
    # build.ps1's own Excel instance has nothing else open.
    close_wb('Builder.xlsm')
    opened.remove('Builder.xlsm')

    # --- run setup ---------------------------------------------------------
    built.run('SetQuiet', True)
    built.run('InitialiseWorkbook')
    built.run('StampBuild')
    built.run('ResetSheetViews')
    log = built.run('QuietLog')
    built.run('SetQuiet', False)
    stamped = built.run('VersionString')
    if not stamped:
        die('VersionString did not answer after setup - the import is incomplete or did not compile.')

    print()
    print('--- InitialiseWorkbook said ---')
    print(log)

    # --- save as .xlsm -----------------------------------------------------
    # filename takes the plain POSIX path; an HFS (colon) path is refused
    # with "Parameter error".
    built._do('''set calculation to calculation manual
set calculate before save to false
save workbook as wb filename (item 2 of argv) file format macro enabled XML file format''',
              work_dst)
    close_wb(os.path.basename(work_dst))
    close_wb(wb_name)
    opened.clear()
    shutil.copy2(work_dst, DST)

    with open(os.path.join(VBA, 'modVersion.bas'), encoding='utf-8') as f:
        ver = re.search(r'APP_VERSION As String = "([^"]+)"', f.read()).group(1)
    ver_name = f'PrintJob-v{ver}-{stamp:%Y%m%d-%H%M}{git_suffix()}.xlsm'
    shutil.copy2(DST, os.path.join(ROOT, ver_name))
    print(f'versioned copy -> {ver_name}')
    keep_newest(os.path.join(ROOT, 'PrintJob-v*.xlsm'), 5)
    print()
    print(f'wrote {DST}')
    print(f'       {stamped}')
finally:
    for n in opened:
        close_wb(n)

# Same safety cap as build.ps1's prune-backups.ps1 -Keep 10.
keep_newest(os.path.join(ROOT, '*.bak.xlsm'), 10)

if os.path.getmtime(SRC) != src_stamp_before:
    print()
    print('WARNING: THE SOURCE .xlsx WAS MODIFIED BY THIS BUILD. That must never happen. '
          'Restore it from OneDrive version history before building again.')

print('done')
