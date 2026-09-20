Attribute VB_Name = "modProtect"
Option Explicit

' Sheet protection.
'
' UserInterfaceOnly is NOT saved with the workbook. A file saved in that state
' reopens fully protected, and VBA can then no longer write to its own sheets -
' which works perfectly in development and fails on every user's machine after
' the first save. Workbook_Open therefore calls ProtectAll every time.
'
' Workbook structure is deliberately left unprotected: users must be able to
' duplicate a location sheet to create a new print room (spec 17).

Public Const PWD As String = "printlog"

Public Sub ProtectAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ProtectSheet ws
    Next ws
End Sub

Public Sub ProtectSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect PWD
    ws.Protect Password:=PWD, UserInterfaceOnly:=True, DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFiltering:=True, AllowSorting:=True, AllowFormattingCells:=False, AllowFormattingColumns:=True, AllowFormattingRows:=True
    ws.EnableOutlining = True
    On Error GoTo 0
End Sub

' Adding or deleting rows in a ListObject is unreliable on a protected sheet
' even with UserInterfaceOnly, so structural work is bracketed by these.
Public Sub UnlockSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect PWD
End Sub

Public Sub RelockSheet(ByVal ws As Worksheet)
    ProtectSheet ws
End Sub
