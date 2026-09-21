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

Public Sub btnExportReport()
    On Error GoTo Fail
    If Not RequireReports Then Exit Sub
    ExportReportSnapshot ActiveSheet
    Exit Sub
Fail:
    ReportError "Export report"
End Sub

Public Sub btnDeleteVisibleReports()
    On Error GoTo Fail
    If Not RequireReports Then Exit Sub
    DeleteVisibleReports
    Exit Sub
Fail:
    ReportError "Delete visible records"
End Sub

Public Sub btnToggleConfigSheets()
    On Error GoTo Fail
    ToggleConfigSheets
    Exit Sub
Fail:
    ReportError "Hide/show settings sheets"
End Sub

Public Sub btnAddRowPrinters()
    On Error GoTo Fail
    AddCatalogRow "tblPrinters"
    Exit Sub
Fail:
    ReportError "Add row"
End Sub

Public Sub btnRemoveRowPrinters()
    On Error GoTo Fail
    RemoveCatalogRow "tblPrinters"
    Exit Sub
Fail:
    ReportError "Remove row"
End Sub

Public Sub btnAddRowPapers()
    On Error GoTo Fail
    AddCatalogRow "tblPapers"
    Exit Sub
Fail:
    ReportError "Add row"
End Sub

Public Sub btnRemoveRowPapers()
    On Error GoTo Fail
    RemoveCatalogRow "tblPapers"
    Exit Sub
Fail:
    ReportError "Remove row"
End Sub

Public Sub btnAddRowTechnicians()
    On Error GoTo Fail
    AddCatalogRow "tblTechnicians"
    Exit Sub
Fail:
    ReportError "Add row"
End Sub

Public Sub btnRemoveRowTechnicians()
    On Error GoTo Fail
    RemoveCatalogRow "tblTechnicians"
    Exit Sub
Fail:
    ReportError "Remove row"
End Sub

Public Sub btnAddRowPaperTypes()
    On Error GoTo Fail
    AddCatalogRow "tblPaperTypes"
    Exit Sub
Fail:
    ReportError "Add row"
End Sub

Public Sub btnRemoveRowPaperTypes()
    On Error GoTo Fail
    RemoveCatalogRow "tblPaperTypes"
    Exit Sub
Fail:
    ReportError "Remove row"
End Sub

' Named Family/Size, not PaperFamilies/StandardSizes - see the DrawSmall
' call in modInit.InitialiseWorkbook for why (Button.Name's 32-character
' limit).
Public Sub btnAddRowFamily()
    On Error GoTo Fail
    AddCatalogRow "tblPaperFamilies"
    Exit Sub
Fail:
    ReportError "Add row"
End Sub

Public Sub btnRemoveRowFamily()
    On Error GoTo Fail
    RemoveCatalogRow "tblPaperFamilies"
    Exit Sub
Fail:
    ReportError "Remove row"
End Sub

Public Sub btnAddRowSize()
    On Error GoTo Fail
    AddCatalogRow "tblStandardSizes"
    Exit Sub
Fail:
    ReportError "Add row"
End Sub

Public Sub btnRemoveRowSize()
    On Error GoTo Fail
    RemoveCatalogRow "tblStandardSizes"
    Exit Sub
Fail:
    ReportError "Remove row"
End Sub

Public Sub btnAddRowConsumable()
    On Error GoTo Fail
    AddCatalogRow "tblConsumables"
    Exit Sub
Fail:
    ReportError "Add row"
End Sub

Public Sub btnRemoveRowConsumable()
    On Error GoTo Fail
    RemoveCatalogRow "tblConsumables"
    Exit Sub
Fail:
    ReportError "Remove row"
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

Private Function RequireReports() As Boolean
    If TypeOf ActiveSheet Is Worksheet Then
        If StrComp(ActiveSheet.Name, "Reports", vbTextCompare) = 0 Then
            RequireReports = True
            Exit Function
        End If
    End If
    Say "This command only works on the Reports sheet.", "It acts on whatever is currently visible there.", "Go to the Reports sheet, then try again."
End Function
