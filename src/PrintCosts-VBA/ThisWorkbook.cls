VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

Private Sub Workbook_Open()
    On Error Resume Next
    AppReset

    ' Not cosmetic: a workbook carrying the 1904 date system would shift every
    ' date by four years, and all date filtering with it.
    If ThisWorkbook.Date1904 Then ThisWorkbook.Date1904 = False

    ' UserInterfaceOnly protection is not saved with the file, so it has to be
    ' re-applied on every open or VBA cannot write to its own sheets.
    ProtectAll
    Invalidate

    ' Repairs the buttons if the file has been renamed since it was saved, and
    ' nothing otherwise.
    HealButtons
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Fail
    If Not TypeOf Sh Is Worksheet Then Exit Sub

    ' Any edit to configuration invalidates the cached catalogue.
    Select Case Sh.Name
        Case "Papers", "Printers", "Print Technicians", "Settings"
            Invalidate
            Exit Sub
    End Select

    If Not IsLocation(Sh) Then Exit Sub
    Application.EnableEvents = False
    OnCellChanged Sh, Target
    Application.EnableEvents = True
    Exit Sub
Fail:
    AppReset
    ReportError "checking that change"
End Sub

' The picker is a worksheet pretending to be a dialog, and it cannot be modal.
' Activating any other sheet therefore counts as cancelling it, so it never
' lingers in the tab list.
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    HideIfLeft Sh
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Fail
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If Not IsLocation(Sh) Then Exit Sub
    OnSelection Sh, Target
    Exit Sub
Fail:
    Application.EnableEvents = True
End Sub

' Double-clicking a blank Date/Time cell stamps it. The Now button does the
' same thing; this is just the shortcut people reach for.
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    On Error GoTo Fail
    Dim lo As ListObject, col As Long
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If Not IsLocation(Sh) Then Exit Sub

    Set lo = JobsTable(Sh)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Application.Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub

    col = Target.Column - lo.Range.Column + 1
    If StrComp(CStr(lo.HeaderRowRange.Cells(1, col).Value), "Date/Time", vbTextCompare) = 0 Then
        Application.EnableEvents = False
        Target.Value = Now
        Application.EnableEvents = True
        Cancel = True
    End If
    Exit Sub
Fail:
    Application.EnableEvents = True
End Sub
