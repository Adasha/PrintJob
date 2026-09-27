# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Shared fixture, dot-sourced by test scripts that need a SECOND, populated
# location sheet ("Annexe", table tblJobs_ANNEX) to exercise workbook-wide
# behaviour: the reduced-view toggle applying to every location sheet, the
# Reports page's bulk delete spanning multiple rooms, Import moving a row
# between rooms, Backup/Restore covering more than one room's job table.
#
# PrintCosts.xlsx stopped shipping a pre-built Annexe sheet 2026-09-26
# (ARCHITECTURE.md's addendum: a second pre-built sheet added nothing the
# documented duplicate-and-rename procedure, SETUP §4.4, didn't already give
# a user, and risked "which one is the real template"). That is the right
# call for the shipped product, but it left every test below assuming a sheet
# that no longer exists. Rather than hand-building a second sheet from
# scratch (which would drift from whatever a real duplicate actually looks
# like), this recreates one the same way a user would - Worksheet.Copy, via
# VBA injected into the caller's own already-open %TEMP% copy, same idiom
# test-duplicate.ps1 already uses for its own injected test code. Nothing
# here is saved, and the deliverable itself is never opened.
#
# Deliberately does NOT follow section 4.4's own Clear All step: that step
# exists for a real new location a user is setting up from scratch and
# doesn't want a copied room's records under its name. These tests want the
# opposite - real job data to exercise - the same reason the original shipped
# Annexe carried a sample row of its own, so the duplicate's rows are left in
# place.
#
# The new sheet's LOC_Code is set to "ANNEX" before Refresh Locations runs,
# rather than left as the copied "MAIN": AssignCode (modRegistry) reads
# whatever LOC_Code already holds first, and Example Print Room reaches the
# refresh loop first (it's earlier in the workbook), so an unchanged "MAIN"
# on the copy would only be deduplicated to something like "MAIN2" - not the
# "ANNEX" every dependent test hardcodes via tblJobs_ANNEX.

function Add-AnnexeFixture($xl, $wb) {
    $vba = @'
Public Function CreateAnnexeFixture() As String
    Dim r As String, src As Worksheet, ws As Worksheet
    Dim codeCell As Range, nameCell As Range
    On Error GoTo Fail
    Application.DisplayAlerts = False

    Set src = ThisWorkbook.Worksheets("Example Print Room")
    src.Copy After:=src
    Set ws = ActiveSheet
    ws.Name = "Annexe"
    r = "created 'Annexe' by duplicating 'Example Print Room' (" & ws.ListObjects(1).ListRows.Count & " rows carried over)" & vbCrLf

    UnlockSheet ws
    Set codeCell = LocRange(ws, "LOC_Code")
    If Not codeCell Is Nothing Then codeCell.Value = "ANNEX"
    Set nameCell = LocRange(ws, "LOC_Name")
    If Not nameCell Is Nothing Then nameCell.Value = "Annexe"
    RelockSheet ws

    SetQuiet True
    RefreshLocations
    r = r & "Refresh Locations said:" & vbCrLf & QuietLog
    SetQuiet False

    CreateAnnexeFixture = r
    Exit Function
Fail:
    CreateAnnexeFixture = r & "ERROR " & Err.Number & ": " & Err.Description & vbCrLf
End Function
'@
    $m = $wb.VBProject.VBComponents.Add(1)
    $m.Name = 'modAnnexeFixture'
    $m.CodeModule.AddFromString($vba)
    $result = $xl.Run('CreateAnnexeFixture')
    # Deliberately NOT removed afterward (unlike a tidier VBComponents.Remove
    # would suggest) - reproduced directly: removing this module immediately
    # after running it left the COM session unstable, and the CALLER's very
    # next Application.Run/property access (e.g. test-phase8.ps1's second
    # InitialiseWorkbook run) failed with RPC_E_CALL_REJECTED. Harmless to
    # leave in place - this workbook is a %TEMP% copy that's never saved.
    if ($result -like '*ERROR*') { throw "Add-AnnexeFixture failed:`n$result" }
    return $result
}
