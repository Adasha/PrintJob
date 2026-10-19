# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Shared by the mac/test-*.py regression scripts: the Mac counterpart of
# TestCommon.ps1 plus the COM calls the Windows scripts make inline. Excel for
# Mac has no COM, so every call here is one osascript run against Excel's
# AppleScript dictionary (/Applications/Microsoft Excel.app/Contents/Resources/
# Excel.sdef).
#
# Kept deliberately narrow - open a copy, run a macro, read/write a cell,
# resolve a name, find a table column, Check - so that a test written against
# it maps line for line onto the .ps1 it ports, and onto VBA if the checks
# ever move into the workbook itself.
#
# Traps, all found the hard way (2026-10-19, Excel for Mac 16.114):
#
#   - "run VB Macro" (capital M) is the application command that runs a macro
#     and returns its result as text. "run VB macro" (lower-case m) is a
#     different, range-level command: it accepts any string and does nothing,
#     no error.
#   - A macro that does not exist ALSO returns empty text with no error. A Sub
#     returns empty text too, so an empty result proves nothing; assert on what
#     the macro did, never on the call succeeding.
#   - Macro names are bare ("btnAddPrintJob") or module-qualified
#     ("modJobs.ClearDefaults"). Workbook-qualified ("Book.xlsm!Proc") is
#     silently not found. Bare names resolve against whatever is open, which is
#     why open_copy refuses to run beside another PrintJob workbook.
#   - Only text, numbers and booleans cross over. A Sub taking a Worksheet
#     (ClearDefaults) cannot be called; call its button wrapper, which uses
#     ActiveSheet, after activate().

import os
import shutil
import subprocess
import sys
import tempfile
import time

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DELIVERABLE = os.path.join(REPO, 'src', 'PrintJob.xlsm')
WB_NAME = 'PrintJob.xlsm'
# Excel for Mac is sandboxed. Opening a copy from the usual $TMPDIR (under
# /var/folders) hung the open AppleEvent for two minutes with nothing opened -
# most likely a Grant File Access prompt. The Office group container would
# avoid that, but macOS stops a terminal writing into another app's
# container. /private/tmp opens without a prompt, so the copies go there.
WORK_ROOT = '/private/tmp/PrintCostsTest'

_any_fail = False


class ExcelError(RuntimeError):
    pass


def osa(script, *args, timeout=600):
    """Run an AppleScript whose 'on run argv' takes args as text. Returns stdout."""
    p = subprocess.run(['osascript', '-', *[str(a) for a in args]],
                       input=script, capture_output=True, text=True, timeout=timeout)
    if p.returncode != 0:
        raise ExcelError(p.stderr.strip())
    return p.stdout.rstrip('\n')


# Every per-workbook script starts the same way. argv item 1 is always the
# workbook name; a script's own arguments start at item 2.
_PRE = '''on run argv
tell application "Microsoft Excel"
set wb to workbook (item 1 of argv)
with timeout of 600 seconds
'''
_POST = '''
end timeout
end tell
end run
'''


def col_letters(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def addr(row, col):
    return f'{col_letters(col)}{row}'


def _parse_ref(ref):
    """'='Example Print Room'!$B$4' -> (4, 2). Top-left cell of the range."""
    cell = ref.split('!')[-1].split(':')[0].replace('$', '')
    letters = ''.join(c for c in cell if c.isalpha())
    digits = ''.join(c for c in cell if c.isdigit())
    col = 0
    for c in letters.upper():
        col = col * 26 + ord(c) - 64
    return int(digits), col


class Workbook:
    def __init__(self, name):
        self.name = name

    def _do(self, body, *args):
        return osa(_PRE + body + _POST, self.name, *args)

    # ------------------------------------------------------------- macros
    def run(self, macro, *args):
        """xl.Run(macro, args...). Args are str, int, float or bool."""
        parts = []
        for i, a in enumerate(args, start=1):
            if isinstance(a, bool):
                parts.append(f'arg{i} {"true" if a else "false"}')
            elif isinstance(a, (int, float)):
                parts.append(f'arg{i} ((item {i + 2} of argv) as number)')
            else:
                parts.append(f'arg{i} (item {i + 2} of argv)')
        body = f'return (run VB Macro (item 2 of argv) {" ".join(parts)})'
        return self._do(body, macro, *['' if isinstance(a, bool) else a for a in args])

    # -------------------------------------------------------------- cells
    def text(self, sheet, a1):
        """Range.Text - what the cell displays."""
        return self._do('return (string value of range (item 3 of argv) of worksheet (item 2 of argv) of wb) as text',
                        sheet, a1)

    def cell_text(self, sheet, row, col):
        return self.text(sheet, addr(row, col))

    def set_value(self, sheet, a1, value):
        """Range.Value2 = value. Fires Worksheet_Change like the COM write does."""
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            v = '((item 4 of argv) as number)'
        else:
            v = '(item 4 of argv)'
        self._do(f'set value2 of range (item 3 of argv) of worksheet (item 2 of argv) of wb to {v}',
                 sheet, a1, value)

    def activate(self, sheet):
        self._do('activate object worksheet (item 2 of argv) of wb', sheet)

    # -------------------------------------------------------------- names
    def name_ref(self, sheet, name):
        """(row, col) a sheet-level defined name refers to, or None if it is missing."""
        try:
            ref = self._do('return references of named item (item 3 of argv) of worksheet (item 2 of argv) of wb',
                           sheet, name)
        except ExcelError:
            return None
        return _parse_ref(ref)

    def name_addr(self, sheet, name):
        r = self.name_ref(sheet, name)
        return addr(*r) if r else None

    # ------------------------------------------------------------- tables
    def table_col(self, sheet, table, header):
        """Worksheet column number of a ListObject column, or 0 if absent
        (TestCommon.ps1's Col, plus the $lo.Range.Column offset every caller adds)."""
        out = self._do('''set lo to list object (item 3 of argv) of worksheet (item 2 of argv) of wb
set c0 to first column index of (range object of lo)
repeat with i from 1 to (count of list columns of lo)
  if (name of list column i of lo) is (item 4 of argv) then return c0 + i - 1
end repeat
return 0''', sheet, table, header)
        return int(out)

    def table_last_row(self, sheet, table):
        """Worksheet row of the table's last data row ($lo.ListRows($lo.ListRows.Count).Range.Row)."""
        out = self._do('''set lo to list object (item 3 of argv) of worksheet (item 2 of argv) of wb
set rg to range object of lo
return (first row index of rg) + (count of rows of rg) - 1''', sheet, table)
        return int(out)


# ------------------------------------------------------------------ copies
class open_copy:
    """with open_copy() as wb: ... drives a COPY of src/PrintJob.xlsm in a temp
    folder, never the deliverable itself, and closes it WITHOUT saving."""

    def __init__(self, deliverable=DELIVERABLE):
        self.deliverable = deliverable
        self.dir = None
        self.wb = None

    def __enter__(self):
        try:
            return self._open()
        except BaseException:
            self.__exit__(None, None, None)
            raise

    def _open(self):
        open_names = osa('''on run argv
tell application "Microsoft Excel"
set out to ""
repeat with w in workbooks
  set out to out & (name of w) & linefeed
end repeat
return out
end tell
end run''').splitlines()
        clash = [n for n in open_names if n.lower().startswith('printjob')]
        if clash:
            raise ExcelError(f'close these first - macro names would resolve against them: {", ".join(clash)}')

        os.makedirs(WORK_ROOT, exist_ok=True)
        self.dir = tempfile.mkdtemp(dir=WORK_ROOT)
        path = os.path.join(self.dir, WB_NAME)
        shutil.copy2(self.deliverable, path)
        osa('''on run argv
tell application "Microsoft Excel"
with timeout of 60 seconds
open POSIX file (item 1 of argv)
end timeout
end tell
end run''', path)
        self.wb = Workbook(WB_NAME)
        # Workbook_Open does real work (ProtectAll, HealButtons, ...). Wait for
        # a function call to come back, which also proves macros are enabled.
        deadline = time.time() + 60
        while True:
            try:
                if self.wb.run('VersionString'):
                    break
            except ExcelError:
                pass
            if time.time() > deadline:
                raise ExcelError('workbook opened but VersionString never answered - are macros disabled?')
            time.sleep(0.5)
        return self.wb

    def __exit__(self, *exc):
        if self.wb:
            try:
                osa('''on run argv
tell application "Microsoft Excel" to close workbook (item 1 of argv) saving no
end run''', self.wb.name)
            except ExcelError as e:
                print(f'  (close failed: {e})')
        print()
        print('closed without saving')
        if self.dir:
            shutil.rmtree(self.dir, ignore_errors=True)
        return False


# ---------------------------------------------------------------- asserting
def check(cond, msg):
    """One assertion line, same format as TestCommon.ps1's Check."""
    global _any_fail
    print(f'  {"OK  " if cond else "FAIL"}  {msg}', flush=True)
    if not cond:
        _any_fail = True


def section(title, first=False):
    if not first:
        print()
    print(f'=== {title} ===', flush=True)


def finish():
    sys.exit(1 if _any_fail else 0)
