Attribute VB_Name = "modProtect"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Sheet protection.
'
' UserInterfaceOnly is NOT saved with the workbook. A file saved in that state
' reopens fully protected, and VBA can then no longer write to its own sheets -
' which works perfectly in development and fails on every user's machine after
' the first save. Workbook_Open therefore calls ProtectAll every time.
'
' Workbook structure is deliberately left unprotected: users must be able to
' duplicate a location sheet to create a new print room (spec 17).
'
' No password (snag list item 7). Sheets are still locked - this only removes
' the PWD that used to gate Unprotect. Add one back by hand in Excel
' (Review > Protect Sheet) if a particular workbook needs it; nothing here
' generates one by default any more.

Public Sub ProtectAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ProtectSheet ws
    Next ws
End Sub

Public Sub ProtectSheet(ByVal ws As Worksheet)
    Dim allowSortFilter As Boolean
    On Error Resume Next
    ' A location sheet is neither sorted nor filtered (multi-pass design decision
    ' 12): rows stay in the order they were added, and a sort would scatter a
    ' job's colour passes from the job. Reports is the place for both.
    allowSortFilter = Not IsLocation(ws)
    ws.Unprotect
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFiltering:=allowSortFilter, AllowSorting:=allowSortFilter, AllowFormattingCells:=False, AllowFormattingColumns:=True, AllowFormattingRows:=True
    ws.EnableOutlining = True
    On Error GoTo 0
End Sub

' Adding or deleting rows in a ListObject is unreliable on a protected sheet
' even with UserInterfaceOnly, so structural work is bracketed by these.
Public Sub UnlockSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect
End Sub

Public Sub RelockSheet(ByVal ws As Worksheet)
    ProtectSheet ws
End Sub
