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
#   - A bare macro name ("btnAddPrintJob") resolves against the ACTIVE
#     workbook only - open a second workbook and the same call silently finds
#     nothing. Workbook.run therefore always qualifies it, quoted:
#     'PrintJob.xlsm'!btnAddPrintJob. The quotes are not optional: unquoted, a
#     name with a hyphen in it (PrintJob-spike.xlsm!Proc) is not found either.
#   - Calling a function that has several Optional parameters, and leaving
#     some out, CRASHES Excel (EXC_BAD_ACCESS in the VBA engine) - two
#     optional parameters left out were fine, eight were not. run() is for
#     macros whose parameters are all given; anything with Optional
#     parameters goes through call(), whose ScriptRun has none.
#   - Inside VBA, Application.Run on a macro name that does not exist also
#     CRASHES Excel for Mac (Windows raises 1004). modScript.ScriptRun checks
#     the name exists first, so a typo in call() is an error, not a crash.
#   - A Boolean a macro returns through run() arrives as AppleScript's
#     'true'/'false'; through call() it is VBA's CStr, 'True'/'False'.
#   - Only text, numbers and booleans cross over. For a macro that takes a
#     Worksheet, Range or ListObject, use call() with sheet()/table()/rng()/
#     spill() tokens: the workbook's modScript.ScriptRun turns them into the
#     objects. An object a call returns comes back as a handle token that a
#     later call can pass in.

import os
import shutil
import subprocess
import tempfile
import sys
import time

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DELIVERABLE = os.path.join(REPO, 'src', 'PrintJob.xlsm')
WB_NAME = 'PrintJob.xlsm'
# Where the throwaway copies (and mac/build.py's work) live. Excel for Mac is
# sandboxed: it can only write where it has been granted access, and VBA
# writing beside a copy (an export, a backup) fails silently - Save As leaves
# an empty file - unless the folder is granted. mac/grant-access.py grants
# these two once.
#
# /Users/Shared because nothing in its path is a link. Under /tmp (a link to
# /private/tmp) Excel reports the workbook's folder as /tmp/..., and from the
# sandbox VBA cannot then read back a file it wrote there - so every "did the
# file get written" check failed. Not $TMPDIR (a sandbox prompt on every
# open), not the Office group container (macOS stops a terminal writing into
# another app's container).
#
# (An empty export and a deleted workbook, first seen here, turned out to be
# Excel's ScreenUpdating bug - modUtils.NewScratchBook - not this folder.)
WORK_ROOT = '/Users/Shared/PrintCostsTest'
BUILD_ROOT = '/Users/Shared/PrintCostsBuild'
GRANT_PATHS = [WORK_ROOT, BUILD_ROOT]

_any_fail = False


class ExcelError(RuntimeError):
    pass


def osa(script, *args, timeout=600, literal=False):
    """Run an AppleScript whose 'on run argv' takes args as text. Returns stdout,
    or with literal=True the result parsed back into Python (lists, str, float,
    bool, None) - plain output flattens a list to 'a, b, c', which a cell
    containing a comma would corrupt."""
    cmd = ['osascript', *(['-ss'] if literal else []), '-', *[str(a) for a in args]]
    p = subprocess.run(cmd, input=script, capture_output=True, text=True, timeout=timeout)
    if p.returncode != 0:
        raise ExcelError(p.stderr.strip())
    out = p.stdout.rstrip('\n')
    return _parse_literal(out) if literal else out


def _parse_literal(s):
    """Parse osascript -ss output: {..} lists, "..." strings, numbers,
    true/false, missing value."""
    pos = 0

    def ws():
        nonlocal pos
        while pos < len(s) and s[pos] in ' \t\r\n':
            pos += 1

    def val():
        nonlocal pos
        ws()
        c = s[pos]
        if c == '{':
            pos += 1
            items = []
            ws()
            if s[pos] == '}':
                pos += 1
                return items
            while True:
                items.append(val())
                ws()
                if s[pos] == ',':
                    pos += 1
                    continue
                pos += 1   # '}'
                return items
        if c == '"':
            pos += 1
            buf = []
            while s[pos] != '"':
                if s[pos] == '\\':
                    pos += 1
                    buf.append({'n': '\n', 'r': '\r', 't': '\t'}.get(s[pos], s[pos]))
                else:
                    buf.append(s[pos])
                pos += 1
            pos += 1
            return ''.join(buf)
        for word, v in (('missing value', None), ('true', True), ('false', False)):
            if s.startswith(word, pos):
                pos += len(word)
                return v
        start = pos
        while pos < len(s) and s[pos] not in ',}':
            pos += 1
        tok = s[start:pos].strip()
        try:
            return float(tok)
        except ValueError:
            return tok   # an enum or anything else unquoted, e.g. validate list

    return val()


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


# Object tokens for Workbook.call - modScript.bas has the list.
WB = '@@wb'


def sheet(name):
    return f'@@sheet:{name}'


def table(name):
    return f'@@table:{name}'


def rng(sheet_name, a1):
    return f'@@range:{sheet_name}!{a1}'


def spill(sheet_name, a1):
    return f'@@spill:{sheet_name}!{a1}'


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

    def _do(self, body, *args, literal=False):
        return osa(_PRE + body + _POST, self.name, *args, literal=literal)

    # ------------------------------------------------------------- macros
    def run(self, macro, *args):
        """xl.Run(macro, args...), always in this workbook. Args are str, int,
        float or bool."""
        macro = "'" + self.name.replace("'", "''") + "'!" + macro
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

    def call(self, macro, *args):
        """$xl.Run(macro, args...) where args may be objects: pass sheet(),
        table(), rng(), spill() or a handle an earlier call() returned.
        Returns the macro's result as text ('' for a Sub), or a handle token
        for an object or array. Raises on an error, and on an empty answer,
        which means ScriptRun itself was not found (a build without modScript)."""
        packed = []
        for a in args:
            if isinstance(a, bool):
                packed.append('b1' if a else 'b0')
            elif isinstance(a, (int, float)):
                packed.append('n' + repr(a))
            else:
                packed.append('s' + str(a))
        out = self.run('ScriptRun', macro, '\x1f'.join(packed))
        if not out:
            raise ExcelError('ScriptRun did not answer - is modScript in this build?')
        if not out.startswith('ok:'):
            raise ExcelError(f'{macro}: {out}')
        return out[3:]

    def prop(self, handle, prop):
        """A property of a handle's object (Name, Address, Count, Rows.Count,
        Column, Row), or of a handle's array ('ubound', 'item r c')."""
        out = self.run('ScriptProp', handle, prop)
        if not out.startswith('ok:'):
            raise ExcelError(f'{handle} {prop}: {out or "ScriptProp did not answer"}')
        return out[3:]

    # -------------------------------------------------------------- cells
    def text(self, sheet, a1):
        """Range.Text - what the cell displays."""
        return self._do('return (string value of range (item 3 of argv) of worksheet (item 2 of argv) of wb) as text',
                        sheet, a1)

    def cell_text(self, sheet, row, col):
        return self.text(sheet, addr(row, col))

    def texts(self, sheet, a1):
        """Range.Text of every cell in a block, as a list of rows of str - one
        call however big the block, where per-cell text() is one each."""
        v = self._do('return string value of range (item 3 of argv) of worksheet (item 2 of argv) of wb',
                     sheet, a1, literal=True)
        if not isinstance(v, list):
            return [['' if v is None else str(v)]]
        if v and not isinstance(v[0], list):
            v = [v]
        return [['' if c is None else str(c) for c in row] for row in v]

    def value(self, sheet, a1):
        """Range.Value2 of one cell: float for a number, str for text, None if empty."""
        v = self._do('return value2 of range (item 3 of argv) of worksheet (item 2 of argv) of wb',
                     sheet, a1, literal=True)
        return None if v == '' else v

    def number(self, sheet, a1):
        """[double]$range.Value2 - an empty cell is 0, as PowerShell's cast makes it."""
        v = self.value(sheet, a1)
        return 0.0 if v is None else float(v)

    def validation_type(self, sheet, a1):
        """Range.Validation.Type as Excel's enum name ('validate list' is COM's 3)."""
        try:
            out = self._do('return validation type of (validation of range (item 3 of argv) of worksheet (item 2 of argv) of wb)',
                           sheet, a1)
        except ExcelError:
            return ''
        return '' if out == 'missing value' else out   # no validation on the cell

    def column_hidden(self, sheet, col):
        out = self._do('return hidden of entire column of range (item 3 of argv) of worksheet (item 2 of argv) of wb',
                       sheet, addr(1, col))
        return out == 'true'

    def calculate_full_rebuild(self):
        self._do('calculate full rebuild')

    def calculate(self):
        """$xl.Calculate()."""
        self._do('calculate')

    def clear_contents(self, sheet, a1):
        self._do('clear contents range (item 3 of argv) of worksheet (item 2 of argv) of wb', sheet, a1)

    def set_column_hidden(self, sheet, col, hidden):
        self._do(f'set hidden of entire column of range (item 3 of argv) of worksheet (item 2 of argv) of wb to {"true" if hidden else "false"}',
                 sheet, addr(1, col))

    def column_lefts(self, sheet, last_col):
        """[Cells(1, c).Left for c in 1..last_col], index 0 = column 1."""
        v = self._do('''set ws to worksheet (item 2 of argv) of wb
set out to {}
repeat with c from 1 to ((item 3 of argv) as integer)
  set end of out to left position of (cell 1 of column c of ws)
end repeat
return out''', sheet, last_col, literal=True)
        return [float(x) for x in v]

    def select(self, sheet, a1):
        """Range.Select, on the active sheet - fires Workbook_SheetSelectionChange."""
        self._do('select range (item 3 of argv) of worksheet (item 2 of argv) of wb', sheet, a1)

    def protect(self, sheet):
        """$ws.Protect() with no arguments."""
        self._do('protect worksheet worksheet (item 2 of argv) of wb', sheet)

    def delete_table_row(self, sheet, table, n):
        """$lo.ListRows(n).Delete() - on an unprotected sheet, as COM needs too."""
        self._do('delete list row ((item 4 of argv) as integer) of list object (item 3 of argv) of worksheet (item 2 of argv) of wb',
                 sheet, table, n)

    def cell_format(self, sheet, a1):
        """The formatting a .ps1 reads off one cell, in COM's terms:
        color = Interior.Color (BGR integer), italic = Font.Italic,
        left_border = Borders(7).LineStyle as Excel's enum name
        ('line style none' is COM's -4142, 'continuous' its 1), locked."""
        v = self._do('''set rg to range (item 3 of argv) of worksheet (item 2 of argv) of wb
set c to color of interior object of rg
set b to line style of (get border rg which border edge left)
return {c, italic of font object of rg, b as text, locked of rg}''', sheet, a1, literal=True)
        rgb, italic, border, locked = v
        r, g, b = (int(x) for x in rgb)
        return {'color': r + g * 256 + b * 65536, 'italic': italic, 'left_border': border, 'locked': locked}

    def row_top_height(self, sheet, row):
        """(Rows(row).Top, Rows(row).Height)."""
        v = self._do('''set r to row ((item 3 of argv) as integer) of worksheet (item 2 of argv) of wb
return {top of r, height of r}''', sheet, row, literal=True)
        return float(v[0]), float(v[1])

    def column_outline_level(self, sheet, col):
        return int(self._do('return outline level of entire column of range (item 3 of argv) of worksheet (item 2 of argv) of wb',
                            sheet, addr(1, col)))

    def dropdown_names(self, sheet):
        """Names of the Form Controls drop-downs on a sheet ($ws.DropDowns())."""
        v = self._do('return name of every dropdown of worksheet (item 2 of argv) of wb', sheet, literal=True)
        if v is None or v == '':
            return []
        return [str(n) for n in v] if isinstance(v, list) else [str(v)]

    def format_conditions(self, sheet, a1):
        """[(Formula1, Font.Color as COM's BGR integer)] of a range's FormatConditions."""
        v = self._do('''set rg to range (item 3 of argv) of worksheet (item 2 of argv) of wb
set out to {}
repeat with i from 1 to (count of format conditions of rg)
  set fc to format condition i of rg
  set end of out to {formula 1 of fc, color of font object of fc}
end repeat
return out''', sheet, a1, literal=True)
        out = []
        for f1, rgb in v:
            r, g, b = (int(x) for x in rgb) if isinstance(rgb, list) else (0, 0, 0)
            out.append((f1, r + g * 256 + b * 65536))
        return out

    def number_format(self, sheet, a1):
        return self._do('return number format of range (item 3 of argv) of worksheet (item 2 of argv) of wb', sheet, a1)

    def validation_formula1(self, sheet, a1):
        """Range.Validation.Formula1 ('' if the cell has none)."""
        try:
            out = self._do('return formula1 of (validation of range (item 3 of argv) of worksheet (item 2 of argv) of wb)',
                           sheet, a1)
        except ExcelError:
            return ''
        return '' if out == 'missing value' else out

    def buttons(self, sheet):
        """Every Form Controls button on a sheet: name, on_action, left, top,
        width, height, text - one call for the lot."""
        v = self._do('''set ws to worksheet (item 2 of argv) of wb
set out to {}
repeat with i from 1 to (count of buttons of ws)
  set bt to button i of ws
  set end of out to {name of bt, on action of bt, left position of bt, top of bt, width of bt, height of bt, caption of bt}
end repeat
return out''', sheet, literal=True)
        keys = ('name', 'on_action', 'left', 'top', 'width', 'height', 'text')
        return [dict(zip(keys, b)) for b in v]

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

    def protected(self, sheet):
        """Worksheet.ProtectContents."""
        return self._do('return protect contents of worksheet (item 2 of argv) of wb', sheet) == 'true'

    def unprotect(self, sheet):
        """Sheets are protected with no password (modProtect), so this needs none."""
        self._do('unprotect worksheet (item 2 of argv) of wb', sheet)

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

    def wb_name(self, name):
        """(sheet, A1) a workbook-level name refers to, or None if it is missing."""
        try:
            ref = self._do('return references of named item (item 2 of argv) of wb', name)
        except ExcelError:
            return None
        sheet, _, cell = ref.lstrip('=').rpartition('!')
        sheet = sheet.strip("'").replace("''", "'")
        return sheet, cell.split(':')[0].replace('$', '')

    def wb_name_text(self, name):
        """$wb.Names.Item(name).RefersToRange.Text."""
        r = self.wb_name(name)
        return self.text(*r) if r else None

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

    def table_headers(self, sheet, table):
        """Every ListColumn name, in order - one call."""
        v = self._do('''set lo to list object (item 3 of argv) of worksheet (item 2 of argv) of wb
set out to {}
repeat with i from 1 to (count of list columns of lo)
  set end of out to name of list column i of lo
end repeat
return out''', sheet, table, literal=True)
        return [str(n) for n in v]

    def table_column_texts(self, sheet, table, header):
        """Text of every data cell in one table column, top to bottom
        (each $lo.DataBodyRange.Cells(i, col)) - one call."""
        col = self.table_col(sheet, table, header)
        first, last = self.table_first_row(sheet, table) + 1, self.table_last_row(sheet, table)
        if col == 0 or last < first:
            return []
        c = col_letters(col)
        return [r[0] for r in self.texts(sheet, f'{c}{first}:{c}{last}')]

    def table_first_row(self, sheet, table):
        """Worksheet row of the table's header ($lo.Range.Row)."""
        return int(self._do('return first row index of (range object of list object (item 3 of argv) of worksheet (item 2 of argv) of wb)',
                            sheet, table))


# Summary table position (modReports SUM_*), as in TestCommon.ps1: the job
# planner box sits above it since 0.10.29. Row of the totals values, of the
# table header, of its first row.
SUM_TOT_ROW, SUM_HDR_ROW, SUM_FIRST_ROW = 24, 27, 28


# ------------------------------------------------------------------ copies
def open_in_excel(path, timeout=90):
    """Open a workbook the way Finder does, and wait until Excel has it.

    Not AppleScript's 'open POSIX file': Excel is sandboxed, and a file named
    that way does not always carry permission to read it - after Excel had
    crashed and relaunched it put up Grant File Access for every copy, and
    the open AppleEvent hung behind the prompt. 'open -a' goes through
    Launch Services, which hands Excel access to the file with the request,
    as a double-click does."""
    name = os.path.basename(path)
    lock = os.path.join(os.path.dirname(path), '~$' + name)
    for attempt in (1, 2):
        subprocess.run(['open', '-a', 'Microsoft Excel', path], check=True)
        # Excel writes its ~$ lock file the moment it starts opening. If 20s
        # pass with neither that nor the workbook, Excel has ignored the
        # request outright - seen after a failed AppleScript sheet copy,
        # with nothing on screen - and only a restart clears it. Done only
        # when Excel has nothing else open, so nothing can be lost.
        if attempt == 1:
            t20 = time.time() + 20
            while time.time() < t20 and not os.path.exists(lock) and name not in open_workbook_names():
                time.sleep(0.5)
            if not os.path.exists(lock) and name not in open_workbook_names() and not open_workbook_names():
                print('  (Excel ignored the open; restarting it)', flush=True)
                osa('tell application "Microsoft Excel" to quit saving no')
                time.sleep(4)
                continue
        break
    deadline = time.time() + timeout
    while name not in open_workbook_names():
        if time.time() > deadline:
            raise ExcelError(f'{name} did not open within {timeout}s - Excel is probably showing a prompt '
                             '(Grant File Access, or a recovery dialog). Answer it, then run this again.')
        time.sleep(0.5)


def open_workbook_names():
    """Names of the workbooks open in Excel. Asked for as one list: walking
    'repeat with w in workbooks' fails with "Parameter error" when Excel has
    reopened after a crash with an unsaved Book1."""
    try:
        v = osa('tell application "Microsoft Excel" to return name of every workbook', literal=True)
    except ExcelError:
        return []   # "Can't get name of every workbook" - none open
    if v is None:
        return []
    return [str(n) for n in v] if isinstance(v, list) else [str(v)]


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
        clash = [n for n in open_workbook_names() if n.lower().startswith('printjob')]
        if clash:
            raise ExcelError(f'close these first - Excel cannot open a second workbook of the same name: {", ".join(clash)}')

        # A fresh folder per run. Reusing one fixed path made Excel say the
        # workbook was locked and offer it read-only - it remembers the
        # document at that path from the run before. open_in_excel grants
        # access per file, so a new folder costs no prompt. Folders from
        # earlier runs are cleared first (no PrintJob workbook is open - just
        # checked).
        os.makedirs(WORK_ROOT, exist_ok=True)
        for old in os.listdir(WORK_ROOT):
            if old.startswith('run-') or old == 'work':
                shutil.rmtree(os.path.join(WORK_ROOT, old), ignore_errors=True)
        self.dir = tempfile.mkdtemp(prefix='run-', dir=WORK_ROOT)
        path = os.path.join(self.dir, WB_NAME)
        shutil.copy2(self.deliverable, path)
        # From here on __exit__ has to close the workbook: even if this open
        # times out (a prompt on screen), Excel finishes opening it later.
        self.wb = Workbook(WB_NAME)
        open_in_excel(path)
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
        closed = True
        if self.wb:
            try:
                osa('''on run argv
tell application "Microsoft Excel" to close workbook (item 1 of argv) saving no
end run''', self.wb.name)
            except ExcelError as e:
                closed = False
                print(f'  (close failed: {e})')
        print()
        print('closed without saving')
        # The copy stays on disk either way; the next run replaces it. A close
        # that fails almost always means Excel crashed or is behind a prompt,
        # and Excel will then reopen the copy by itself - if it had been
        # deleted, it would sit on a Grant File Access prompt for a file that
        # is not there, hanging every script after it.
        if not closed:
            print(f'  (Excel may reopen {self.dir}/{WB_NAME}: close it without saving)')
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
