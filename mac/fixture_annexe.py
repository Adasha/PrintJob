# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-fixture-annexe.ps1: a SECOND, populated location sheet
# ("Annexe", table tblJobs_ANNEX) for the tests that exercise workbook-wide
# behaviour. See that script for why the sheet is made by duplicating Example
# Print Room the way a user would, why its copied rows are kept, and why
# LOC_Code is set to ANNEX before Refresh Locations runs.
#
# Same steps, different route. The .ps1 injects a VBA module and runs it. That
# cannot be done here: AppleScript has no way into the VBA project, so every
# step goes through Excel's AppleScript dictionary instead -
#
#   Worksheet.Copy After:=src         copy worksheet ... after ...
#   ws.Name = "Annexe"                set name of the new sheet
#   UnlockSheet ws                    unprotect (the sheets carry no password)
#   LocRange(...).Value = ...         the copy's sheet-level LOC_Code/LOC_Name
#   SetQuiet / RefreshLocations       run VB Macro, as the .ps1's module does
#
# RelockSheet has no counterpart: RefreshLocations re-protects every location
# sheet itself, and add_annexe_fixture checks that it did.

from xlmac import ExcelError

SRC = 'Example Print Room'
NAME = 'Annexe'


def add_annexe_fixture(wb):
    """Add-AnnexeFixture. Returns the same kind of report the .ps1's does;
    raises ExcelError where that one throws."""
    # The copy's table is auto-renamed by Excel until Refresh Locations names
    # it tblJobs_ANNEX, so it is counted as the sheet's first table, as the
    # .ps1's ws.ListObjects(1) does.
    out = wb._do('''set src to worksheet (item 2 of argv) of wb
copy worksheet src after src
set ws to active sheet
set name of ws to (item 3 of argv)
return (name of ws) & "|" & (count of list rows of list object 1 of ws)''', SRC, NAME)
    new, _, rows = out.rpartition('|')
    if new != NAME:
        raise ExcelError(f"Add-AnnexeFixture failed: the copy is named '{new}', not '{NAME}'")
    r = f"created '{NAME}' by duplicating '{SRC}' ({rows} rows carried over)\n"

    wb.unprotect(NAME)
    for nm, v in (('LOC_Code', 'ANNEX'), ('LOC_Name', NAME)):
        a = wb.name_addr(NAME, nm)
        if a:
            wb.set_value(NAME, a, v)

    wb.run('SetQuiet', True)
    wb.run('RefreshLocations')
    r += 'Refresh Locations said:\n' + wb.run('QuietLog')
    wb.run('SetQuiet', False)

    if not wb.protected(NAME):
        raise ExcelError(f'Add-AnnexeFixture failed: Refresh Locations left {NAME} unprotected\n{r}')
    return r
