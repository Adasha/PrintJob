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

    ' Default to Summary sheet with cursor at column A
    On Error Resume Next
    ThisWorkbook.Sheets("Summary").Activate
    ThisWorkbook.Sheets("Summary").Range("A1").Select
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Fail
    If Not TypeOf Sh Is Worksheet Then Exit Sub

    ' Any edit to configuration invalidates the cached catalogue.
    Select Case Sh.Name
        Case "Summary"
            ' The job planner box: re-price and re-bind its lists. Events off - it writes cells.
            Application.EnableEvents = False
            PlannerChanged Sh, Target
            Application.EnableEvents = True
            Exit Sub
        Case "Papers", "Printers", "Print Technicians", "Departments", "Consumables", "Settings"
            Invalidate
            ' Fills Width/Height from a chosen standard size and defaults
            ' Active - see modCatalog.OnPaperEdited. Events off: it writes cells.
            If Sh.Name = "Papers" Then
                Application.EnableEvents = False
                OnPaperEdited Target
                Application.EnableEvents = True
            ElseIf Sh.Name = "Printers" Then
                ' A row typed under the table gets its ID here - see
                ' modCatalog "catalogue IDs". Events off: it writes cells.
                Application.EnableEvents = False
                OnCatalogEdited "tblPrinters", Target
                Application.EnableEvents = True
            ElseIf Sh.Name = "Print Technicians" Then
                Application.EnableEvents = False
                OnCatalogEdited "tblTechnicians", Target
                Application.EnableEvents = True
            ElseIf Sh.Name = "Departments" Then
                ' ID, Active/Free defaults, tidy aliases, clash warning - modDepartments.
                Application.EnableEvents = False
                OnDepartmentEdited Target
                Application.EnableEvents = True
            ElseIf Sh.Name = "Consumables" Then
                ' ID, Active default, duplicate-name warning - modConsumables.
                Application.EnableEvents = False
                OnColourEdited Target
                Application.EnableEvents = True
            End If
            Exit Sub
    End Select

    If Not IsLocation(Sh) Then Exit Sub
    Application.EnableEvents = False
    OnCellChanged Sh, Target
    RefreshStatusNotes Sh
    Application.EnableEvents = True
    Exit Sub
Fail:
    AppReset
    ReportError "checking that change"
End Sub

' The picker is a worksheet pretending to be a dialog, and it cannot be modal.
' Activating any other sheet therefore counts as cancelling it, so it never
' lingers in the tab list.
'
' Arriving on Reports also refreshes its own Technician/Printer/Paper Stock/
' Location filter dropdowns (modReports.RefreshReportFilterLists). Technician
' there is sourced from what has actually been recorded on _Data, not the
' catalogue (see that sub's own comment) - so unlike the catalogue-driven
' dropdowns above, it goes stale from a job being added, removed or cleared
' on ANY print room, not from a catalogue edit. Rather than chase every place
' a job row can change, refreshing on arrival guarantees it is current
' whenever it is actually looked at.
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    HideIfLeft Sh
    If TypeOf Sh Is Worksheet Then
        If Sh.Name = "Reports" Then
            RefreshReportFilterLists Sh
            RepositionReportsButtons Sh
        End If
        ' Prices or print rooms may have changed since the planner was last looked at.
        If Sh.Name = "Summary" Then RefreshPlanner Sh
        ' Catches a location sheet's buttons drifting from a manually
        ' collapsed/expanded cost-columns group left over from the last visit
        ' - see modInit.RepositionLocationButtons' own comment.
        If IsLocation(Sh) Then RepositionLocationButtons Sh
    End If
End Sub

' Rebinds every print room's Printer/Paper Stock/Technician dropdown the
' moment the user leaves a catalogue sheet they actually edited - not on
' every keystroke inside the edit (Workbook_SheetChange already fires once
' per cell, and thrashing a full rebind that often would make editing
' tblPrinters/tblPapers/tblTechnicians feel sluggish), and not on every sheet
' switch either (CatalogDirty is False the rest of the time, so merely
' looking at Printers/Papers/Print Technicians/Settings costs nothing).
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    On Error GoTo Fail
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If Not CatalogDirty() Then Exit Sub

    Select Case Sh.Name
        Case "Papers", "Printers", "Print Technicians", "Departments", "Consumables", "Settings"
            AppOff
            RebindAllLocationDropdowns
            ClearCatalogDirty
            AppOn
    End Select
    Exit Sub
Fail:
    AppReset
    ReportError "refreshing dropdowns"
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Fail
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    ' Reports: keep the header buttons where their columns are now.
    If StrComp(Sh.Name, "Reports", vbTextCompare) = 0 Then
        RepositionReportsButtons Sh
        Exit Sub
    End If
    If Not IsLocation(Sh) Then Exit Sub
    OnSelection Sh, Target
    ' Same self-heal as Workbook_SheetActivate, for the case where the user
    ' hides/unhides columns by hand (or collapses/expands the cost-columns
    ' group) without ever leaving the sheet - the next click anywhere on it
    ' re-settles the buttons. Needed: without it the side panel stays at
    ' column 36's old position (test-manualhide.ps1). Cheap
    ' (a handful of shape/lookup ops against the table's ~30 headers, not the
    ' data rows) and idempotent, so running it on every selection is fine.
    RepositionLocationButtons Sh
    ' Keeps each problem row's hover note in step with its Status text - see
    ' modInit.RefreshStatusNotes.
    RefreshStatusNotes Sh
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
