Attribute VB_Name = "modInit"
Option Explicit

' One-time setup, run once after the modules are imported.
'
' The workbook file is written outside Excel, which can lay out cells but
' cannot draw controls. This turns the placeholder labels into Form Control
' buttons, applies protection, and leaves the file ready to save as .xlsm.
'
' It is safe to run again at any time - it removes the buttons it created
' before redrawing them - and must be run again after a new print room sheet
' is duplicated.

Private Const BTN_TAG As String = "pcb_"

' Set by a setup run and consumed by the message Refresh Locations shows, so a
' full setup can say how many buttons it drew without a second dialog.
Public gButtonsDrawn As Long

Public Sub InitialiseWorkbook()
    On Error GoTo Fail
    Dim ws As Worksheet, n As Long

    AppOff
    If ThisWorkbook.Date1904 Then ThisWorkbook.Date1904 = False

    ' _Registry, _Audit and _Data are built here rather than shipped in the
    ' .xlsx, so the workbook file stays something VBA can reconstruct.
    EnsureSystemSheets

    ' The version rows and the About block likewise. Setup does not stamp the
    ' build date - that is build.ps1's job, via StampBuild.
    EnsureVersionSettings
    WriteAbout

    ' Summary and Reports, likewise built here. Their formulas read
    ' _Data, which RefreshLocations writes at the end of this run - until then
    ' they sit on their IFERROR fallbacks rather than showing errors.
    BuildReportSheets

    For Each ws In ThisWorkbook.Worksheets
        UnlockSheet ws
        ClearButtons ws
        If IsLocation(ws) Then
            DrawLocationButtons ws
            ConfigValidation ws
            BindColumns ws
            n = n + 1
        ElseIf StrComp(ws.Name, "Summary", vbTextCompare) = 0 Then
            ' Column O onwards, clear of the A:M report table.
            DrawOne ws, 1, 15, "Refresh Locations", "btnRefreshLocations", 130
            DrawOne ws, 3, 15, "Check workbook", "btnCheckWorkbook", 130
            DrawOne ws, 5, 15, "Go to Settings", "btnGoSettings", 130
        ElseIf StrComp(ws.Name, "Printers", vbTextCompare) = 0 Then
            DrawOne ws, 10, 8, "Select families...", "btnSelectFamilies", 130
        ElseIf StrComp(ws.Name, "Settings", vbTextCompare) = 0 Then
            DrawOne ws, 2, 20, "Refresh Locations", "btnRefreshLocations", 130
            DrawOne ws, 4, 20, "Check workbook", "btnCheckWorkbook", 130
            DrawOne ws, 6, 20, "Re-stamp prices...", "btnReStamp", 130
            DrawOne ws, 8, 20, "About", "btnAbout", 130
            DrawOne ws, 10, 20, "Export All Locations...", "btnExportAll", 130
            DrawOne ws, 12, 20, "Import (choose room)...", "btnImportGlobal", 130
        ElseIf StrComp(ws.Name, "Reports", vbTextCompare) = 0 Then
            ' Rows 1-3, column F: clear of the title text (A1:A2) and above
            ' the filter/sort boxes (rows 5+), so both buttons sit inside the
            ' first screenful on any normal window - no scrolling needed.
            DrawOne ws, 1, 6, "Export report...", "btnExportReport", 140
            DrawOne ws, 3, 6, "Delete visible records...", "btnDeleteVisibleReports", 140
        End If
    Next ws

    Invalidate
    ProtectAll
    AppOn

    ' Refresh does the rest - codes, table names, the registry and the
    ' reporting range - and reports on the whole run, so there is one message
    ' at the end rather than two.
    gButtonsDrawn = CountButtons
    RefreshLocations
    Exit Sub
Fail:
    AppReset
    ReportError "InitialiseWorkbook"
End Sub

Private Sub DrawLocationButtons(ByVal ws As Worksheet)
    ' The placeholder labels sit on row 10, two columns apart. Buttons are drawn
    ' over them and the labels cleared.
    DrawOne ws, 10, 1, "Add Print Job", "btnAddPrintJob", 110
    DrawOne ws, 10, 3, "Now", "btnNow", 110
    DrawOne ws, 10, 5, "Remove Row", "btnRemoveRow", 110
    DrawOne ws, 10, 7, "Select printers...", "btnSelectPrinters", 110
    DrawOne ws, 10, 9, "Check this sheet", "btnCheckSheet", 110
    DrawOne ws, 10, 11, "Clear All", "btnClearAll", 110
    DrawOne ws, 10, 13, "Export...", "btnExport", 110
    DrawOne ws, 10, 15, "Import...", "btnImportLocation", 110

    Dim c As Long
    For c = 1 To 15
        With ws.Cells(10, c)
            .ClearContents
            .Interior.Pattern = xlNone
        End With
    Next c
    ws.Cells(10, 13).ClearContents
End Sub

' Yes/No validation on the two location defaults (spec 9.2). These seed each
' new print job and are plain cells in the workbook file, so the list is added
' here rather than shipped with it.
Private Sub ConfigValidation(ByVal ws As Worksheet)
    AddYesNo LocRange(ws, "LOC_DefDisPaper"), "Disregard paper cost", "Sets what new print jobs on this sheet start with. Changing it never alters jobs already recorded."
    AddYesNo LocRange(ws, "LOC_DefDisCons"), "Disregard consumable cost", "Sets what new print jobs on this sheet start with. Changing it never alters jobs already recorded."
End Sub

Private Sub AddYesNo(ByVal target As Range, ByVal Title As String, ByVal Msg As String)
    If target Is Nothing Then Exit Sub
    If Len(Trim$(CStr(target.Value))) = 0 Then target.Value = "No"
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = "Choose Yes or No."
    End With
End Sub

Private Function CountButtons() As Long
    Dim ws As Worksheet, i As Long, t As Long
    For Each ws In ThisWorkbook.Worksheets
        For i = 1 To ws.Buttons.Count
            If Left$(ws.Buttons(i).Name, Len(BTN_TAG)) = BTN_TAG Then t = t + 1
        Next i
    Next ws
    CountButtons = t
End Function

Public Sub DrawOne(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal ColNo As Long, ByVal Caption As String, ByVal Macro As String, ByVal W As Single)
    Dim b As Button, c As Range
    Set c = ws.Cells(RowNo, ColNo)
    Set b = ws.Buttons.Add(c.Left, c.Top, W, 22)
    b.Name = BTN_TAG & Macro & "_" & RowNo & "_" & ColNo
    b.Caption = Caption
    b.OnAction = Macro
    b.Characters.Font.Size = 10
End Sub

Private Sub ClearButtons(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Buttons.Count To 1 Step -1
        If Left$(ws.Buttons(i).Name, Len(BTN_TAG)) = BTN_TAG Then ws.Buttons(i).Delete
    Next i
End Sub
