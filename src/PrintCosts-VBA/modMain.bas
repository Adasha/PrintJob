Attribute VB_Name = "modMain"
Option Explicit

' Entry points bound to the on-sheet buttons.
'
' These names are a contract with the Form Controls: renaming one silently
' breaks a button. Each is a thin wrapper whose only jobs are to work out which
' sheet it is on and to guarantee that application state is restored even when
' something fails.

Public Sub btnAddPrintJob()
    On Error GoTo Fail
    If Not RequireLocation Then Exit Sub
    AddPrintJob ActiveSheet
    Exit Sub
Fail:
    ReportError "Add Print Job"
End Sub

Public Sub btnNow()
    On Error GoTo Fail
    If Not RequireLocation Then Exit Sub
    StampNow ActiveSheet
    Exit Sub
Fail:
    ReportError "Now"
End Sub

Public Sub btnRemoveRow()
    On Error GoTo Fail
    If Not RequireLocation Then Exit Sub
    RemoveRow ActiveSheet
    Exit Sub
Fail:
    ReportError "Remove Row"
End Sub

Public Sub btnExport()
    On Error GoTo Fail
    If Not RequireLocation Then Exit Sub
    ExportLocation ActiveSheet
    Exit Sub
Fail:
    ReportError "Export"
End Sub

Public Sub btnExportAll()
    On Error GoTo Fail
    ExportAllLocations
    Exit Sub
Fail:
    ReportError "Export All Locations"
End Sub

Public Sub btnImportLocation()
    On Error GoTo Fail
    If Not RequireLocation Then Exit Sub
    ImportIntoLocation ActiveSheet
    Exit Sub
Fail:
    ReportError "Import"
End Sub

Public Sub btnImportGlobal()
    On Error GoTo Fail
    ImportGlobal
    Exit Sub
Fail:
    ReportError "Import"
End Sub

Public Sub btnClearAll()
    On Error GoTo Fail
    If Not RequireLocation Then Exit Sub
    ClearAll ActiveSheet
    Exit Sub
Fail:
    ReportError "Clear All"
End Sub

Public Sub btnCheckSheet()
    On Error GoTo Fail
    If Not RequireLocation Then Exit Sub
    LoadCatalog True
    CheckSheet ActiveSheet
    Exit Sub
Fail:
    ReportError "Check this sheet"
End Sub

Public Sub btnGoSettings()
    On Error Resume Next
    ThisWorkbook.Worksheets("Settings").Activate
End Sub

Public Sub btnAbout()
    On Error GoTo Fail
    ShowAbout
    Exit Sub
Fail:
    ReportError "About"
End Sub

Public Sub btnRefreshLocations()
    On Error GoTo Fail
    RefreshLocations
    Exit Sub
Fail:
    ReportError "Refresh Locations"
End Sub

Public Sub btnCheckWorkbook()
    On Error GoTo Fail
    LoadCatalog True
    CheckWorkbook
    Exit Sub
Fail:
    ReportError "Check workbook"
End Sub

Public Sub btnSelectPrinters()
    On Error GoTo Fail
    If Not RequireLocation Then Exit Sub
    PickPrinters ActiveSheet
    Exit Sub
Fail:
    ReportError "Select printers"
End Sub

Public Sub btnSelectFamilies()
    On Error GoTo Fail
    PickFamilies
    Exit Sub
Fail:
    ReportError "Select families"
End Sub

Public Sub btnReStamp()
    On Error GoTo Fail
    LoadCatalog True
    ReStampAll
    Exit Sub
Fail:
    ReportError "Re-stamp prices"
End Sub

Private Function RequireLocation() As Boolean
    If TypeOf ActiveSheet Is Worksheet Then
        If IsLocation(ActiveSheet) Then
            RequireLocation = True
            Exit Function
        End If
    End If
    Say "This command only works on a print room sheet.", "The sheet you are on is not set up as a print room.", "Switch to a print room tab and try again."
End Function
