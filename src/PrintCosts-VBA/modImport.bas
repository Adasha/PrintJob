Attribute VB_Name = "modImport"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Restoring a backup, moving jobs between rooms, or pulling several rooms'
' exports into one copy for reporting - all the same mechanism (snag list
' Problem 1). Reads a file modExport wrote and merges it into one location's
' job table by Job ID.
'
' No cost reconciliation and no catalogue reconciliation. A job row's formulas
' only ever read its own snapshot columns and inputs, never tblPapers or
' tblPrinters, so an imported row already carries the rates it needs and costs
' correctly regardless of what the target workbook's config tables contain.
' Technician/Printer/Paper Stock values are carried as plain text even when
' absent from the target's current lists - import never creates, matches or
' edits a row in tblPrinters/tblTechnicians/tblPapers/tblConsumables.
'
' Conflict handling is by Job ID (globally unique, never reassigned - see
' modRegistry.NextJobId). No existing row with that ID: append. An existing
' row: overwrite it with the imported version. Overwrite, not skip, is
' deliberate - it is what makes "restore a backup over stale/edited data" and
' "re-aggregate overlapping exports" both work correctly.
'
' Only the input and snapshot columns are written - Job ID, Date/Time,
' Student Name, Student No, Technician, Printer, Paper Stock, Unit, Qty,
' Print Width mm, Sheet size, Disregard Paper, Disregard Consumable, Paid,
' Notes, and every S_* column. The calculated columns (Area m2, Paper Cost, Consumable
' Cost, Gross Cost, Disregarded, Chargeable Cost) and Status are left alone -
' the job table's own per-row formulas reproduce the historical figures
' exactly from the snapshot, the same way a locally entered job does.
'
' A file exported before 2026-09-22 (schema 1.0) has no Paid column at all -
' WriteImportedRow reads it the same way as every other field and gets back
' Empty, which writes as blank, which every consumer already treats as "not
' paid" (modUtils.SCHEMA_VER's 1.1 comment). No migration step needed.
'
' Same reasoning covers the Quantity->Qty header rename (0.9.14): a file
' exported before the rename has a column literally called "Quantity", so
' the dict `d` below (keyed by whatever header text the FILE itself uses)
' has no "Qty" entry - WriteNum's d.Item("Qty") comes back Empty and writes
' blank, exactly the same graceful-degradation path as a missing Paid
' column, not a hard failure. Nothing here bulk-migrates old export files.
'
' All writes happen with EnableEvents False, same as every other write this
' workbook makes to its own sheets, so imported rows bypass Worksheet_Change
' on purpose: StampRow must never re-derive S_* from the target's live
' catalogue, or the whole point of the snapshot - costing history at the
' prices it was actually entered at - would be lost the moment it is
' imported into a workbook with different current prices.

Private Const HDR_ROW As Long = 9
Private Const FIRST_DATA_ROW As Long = 10

' ============================================================ entry points ===
Public Sub ImportIntoLocation(ByVal ws As Worksheet)
    Dim path As String, rows As Collection
    On Error GoTo Fail
    path = PickImportFile()
    If Len(path) = 0 Then Exit Sub
    Set rows = ReadImportRows(path)
    ApplyImport ws, rows
    Exit Sub
Fail:
    AppReset
    ReportError "Import"
End Sub

Public Sub ImportGlobal()
    Dim ws As Worksheet, path As String, rows As Collection
    On Error GoTo Fail
    Set ws = PickTargetLocation()
    If ws Is Nothing Then Exit Sub
    path = PickImportFile()
    If Len(path) = 0 Then Exit Sub
    Set rows = ReadImportRows(path)
    ApplyImport ws, rows
    Exit Sub
Fail:
    AppReset
    ReportError "Import"
End Sub

Private Function PickTargetLocation() As Worksheet
    Dim ws As Worksheet, list As String, code As String, answer As Variant

    For Each ws In LocationSheets()
        list = list & "- " & LocValue(ws, "LOC_Code") & "  (" & LocValue(ws, "LOC_Name") & ")" & vbCrLf
    Next ws
    If Len(list) = 0 Then
        Say "No print rooms were found.", "Run Refresh Locations first."
        Exit Function
    End If

    Do
        answer = Application.InputBox( _
            "Which print room should this import go into? Type its code." & vbCrLf & vbCrLf & list, _
            "Import - choose a print room", Type:=2)
        If VarType(answer) = vbBoolean Then
            If answer = False Then Exit Function   ' Cancel
        End If
        code = Trim$(CStr(answer))
        If Len(code) > 0 Then
            Set ws = SheetForCode(code)
            If Not ws Is Nothing Then
                Set PickTargetLocation = ws
                Exit Function
            End If
            Say "'" & code & "' is not one of the print rooms listed.", "", "Check the code and try again."
        End If
    Loop
End Function

Private Function PickImportFile() As String
    PickImportFile = PickCsvFile("Import print jobs - choose a file")
End Function

' ================================================================= merge ===
' Counts the change, asks, and hands off to ApplyImportConfirmed. Split out so
' a test can drive ApplyImportConfirmed directly - Ask() always declines under
' SetQuiet (an unattended run must never confirm a destructive op on the
' user's behalf), same reason modJobs.RemoveRow/ClearAll are shaped this way.
Public Sub ApplyImport(ByVal ws As Worksheet, ByVal rows As Collection)
    Dim lo As ListObject, appended As Long, overwrite As Long, d1 As clsDict
    Set lo = JobsTable(ws)
    If lo Is Nothing Then
        Say "This sheet has no print job table.", "Import only works on a print room sheet."
        Exit Sub
    End If
    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then
        Say "There is nothing to import.", "The file has no data rows."
        Exit Sub
    End If

    CountChange lo, rows, appended, overwrite

    Dim noNames As String
    Set d1 = rows.Item(1)
    If Not d1.Exists("Student Name") Then _
        noNames = vbCrLf & "This file has no student names; names on overwritten records are kept." & vbCrLf
    If Not Ask("Import into '" & LocValue(ws, "LOC_Name") & "'?" & vbCrLf & vbCrLf & _
        appended & " new record" & IIf(appended = 1, "", "s") & " will be added." & vbCrLf & _
        overwrite & " existing record" & IIf(overwrite = 1, "", "s") & " will be OVERWRITTEN with the imported version." & vbCrLf & noNames & vbCrLf & _
        "This cannot be undone.", "Import") Then Exit Sub

    ApplyImportConfirmed ws, rows
End Sub

' The actual write. Public so a test can call it directly, bypassing the
' Ask() gate above the same way test scripts already poke ListObjects
' directly rather than driving a button that would need a confirm.
Public Sub ApplyImportConfirmed(ByVal ws As Worksheet, ByVal rows As Collection)
    Dim lo As ListObject, d As clsDict, jobId As String, n As Long
    Dim appended As Long, overwrite As Long, v As Variant

    On Error GoTo Fail
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    AppOff
    UnlockSheet ws
    For Each v In rows
        Set d = v
        jobId = Trim$(CStr(d.Item("Job ID")))
        If Len(jobId) > 0 Then
            n = FindRowByJobId(lo, jobId)
            If n = 0 Then
                appended = appended + 1
                If lo.ListRows.Count = 1 And IsBlankRow(lo, 1) Then
                    n = 1
                Else
                    n = lo.ListRows.Add.Index
                End If
            Else
                overwrite = overwrite + 1
            End If
            WriteImportedRow lo, n, d
        End If
    Next v
    RelockSheet ws
    Application.Calculate
    AppOn

    LogAudit "Import", LocValue(ws, "LOC_Name"), appended & " added, " & overwrite & " overwritten"

    ' Imported rows bypass Worksheet_Change, so Status/H_Issues need a sweep
    ' rather than relying on the event that normally drives it. CheckSheet
    ' both forces that and tells the user what it found - the more thorough
    ' of the two options the design discussion left open.
    CheckSheet ws
    Exit Sub
Fail:
    AppReset
    ReportError "Import"
End Sub

Private Sub CountChange(ByVal lo As ListObject, ByVal rows As Collection, ByRef Appended As Long, ByRef Overwrite As Long)
    Dim d As clsDict, jobId As String, v As Variant
    Appended = 0: Overwrite = 0
    For Each v In rows
        Set d = v
        jobId = Trim$(CStr(d.Item("Job ID")))
        If Len(jobId) > 0 Then
            If FindRowByJobId(lo, jobId) > 0 Then
                Overwrite = Overwrite + 1
            Else
                Appended = Appended + 1
            End If
        End If
    Next v
End Sub

Private Function FindRowByJobId(ByVal lo As ListObject, ByVal JobId As String) As Long
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Job ID").Value)), JobId, vbTextCompare) = 0 Then
            FindRowByJobId = i
            Exit Function
        End If
    Next i
End Function

Private Sub WriteImportedRow(ByVal lo As ListObject, ByVal RowNo As Long, ByVal d As clsDict)
    WriteText lo, RowNo, "Job ID", d
    WriteDate lo, RowNo, "Date/Time", d
    If d.Exists("Student Name") Then WriteText lo, RowNo, "Student Name", d
    WriteText lo, RowNo, "Student No", d
    WriteText lo, RowNo, "Technician", d
    WriteText lo, RowNo, "Printer", d
    WriteText lo, RowNo, "Paper Stock", d
    ' Unit is a calculated column (it now follows the room's roll length
    ' unit), so it is not written - that would freeze the formula to a
    ' string for this row. The file's own Unit text is only read, below, to
    ' say what unit its Qty is in: a roll length is converted if the file
    ' (cm, or metres/older files) and this room's roll length unit differ.
    WriteNum lo, RowNo, "Qty", d
    ConvertImportedRollQty lo, RowNo, Trim$(CStr(d.Item("Unit")))
    WriteNum lo, RowNo, "Print Width mm", d
    ' Missing from a file exported before this column existed (schema < 1.2)
    ' - d.Item returns Empty, which writes blank, same graceful degradation
    ' as Paid's own comment above.
    WriteText lo, RowNo, "Sheet size", d
    WriteText lo, RowNo, "Disregard Paper", d
    WriteText lo, RowNo, "Disregard Consumable", d
    ' Missing from a file exported before 2026-09-22 (schema 1.0) - d.Item
    ' returns Empty for a key the source file never had, so this writes blank
    ' rather than erroring, and blank reads as "not paid" everywhere else
    ' (Summary/Reports totals).
    WriteText lo, RowNo, "Paid", d
    WriteText lo, RowNo, "Notes", d
    WriteText lo, RowNo, "S_PrinterID", d
    WriteText lo, RowNo, "S_StockID", d
    WriteText lo, RowNo, "S_TechID", d
    WriteText lo, RowNo, "S_Measure", d
    WriteNum lo, RowNo, "S_UnitCost", d
    WriteNum lo, RowNo, "S_StockWidth_mm", d
    WriteNum lo, RowNo, "S_SheetHeight_mm", d
    WriteNum lo, RowNo, "S_ConsRate", d
    WriteDate lo, RowNo, "S_StampedAt", d
    WriteText lo, RowNo, "S_StampedBy", d
    WriteText lo, RowNo, "S_SchemaVer", d
End Sub

Private Sub ConvertImportedRollQty(ByVal lo As ListObject, ByVal RowNo As Long, ByVal FileUnit As String)
    Dim c As Range, fileCm As Boolean, roomCm As Boolean
    If StrComp(FileUnit, "cm", vbTextCompare) <> 0 And StrComp(FileUnit, "metres", vbTextCompare) <> 0 Then Exit Sub
    Set c = CellIn(lo, RowNo, "Qty")
    If IsEmpty(c.Value2) Then Exit Sub
    fileCm = (StrComp(FileUnit, "cm", vbTextCompare) = 0)
    roomCm = (RollUnitOf(lo.Parent) = "Centimetres")
    If roomCm And Not fileCm Then c.Value2 = Round(CDbl(c.Value2) * 100, 6)
    If fileCm And Not roomCm Then c.Value2 = Round(CDbl(c.Value2) / 100, 6)
End Sub

Private Sub WriteText(ByVal lo As ListObject, ByVal RowNo As Long, ByVal Header As String, ByVal d As clsDict)
    Dim s As String
    s = Trim$(CStr(d.Item(Header)))
    If Len(s) = 0 Then
        CellIn(lo, RowNo, Header).ClearContents
    Else
        CellIn(lo, RowNo, Header).Value = s
    End If
End Sub

' A real Double, not a text digit-string, so downstream formulas that add,
' compare or feed these into modUtils.NumOf behave exactly as they do for a
' locally entered job. CDbl is locale-aware, matching how the rest of this
' project reads a numeric string (see modUtils.NumOf's own comment on why
' Val() is the wrong tool here).
Private Sub WriteNum(ByVal lo As ListObject, ByVal RowNo As Long, ByVal Header As String, ByVal d As clsDict)
    Dim s As String
    s = Trim$(CStr(d.Item(Header)))
    If Len(s) = 0 Or Not IsNumeric(s) Then
        CellIn(lo, RowNo, Header).ClearContents
    Else
        CellIn(lo, RowNo, Header).Value = CDbl(s)
    End If
End Sub

Private Sub WriteDate(ByVal lo As ListObject, ByVal RowNo As Long, ByVal Header As String, ByVal d As clsDict)
    Dim s As String
    s = Trim$(CStr(d.Item(Header)))
    If Len(s) = 0 Then
        CellIn(lo, RowNo, Header).ClearContents
    Else
        CellIn(lo, RowNo, Header).Value = ParseIsoDateTime(s)
    End If
End Sub

' ================================================================== read ===
' Reads a file modExport wrote back into memory: one clsDict per data row,
' keyed by header text, so a column can be reordered (or a future schema can
' add or drop one) without this breaking - the same freedom modExport's own
' comment claims for column order on the way out.
'
' Public so a test can drive the read half of import directly - PickImportFile
' opens a real OS file picker that cannot run unattended.
Public Function ReadImportRows(ByVal path As String) As Collection
    Dim wbIn As Workbook, ws As Worksheet
    Dim fi() As Variant, i As Long, c As Long, r As Long
    Dim lastRow As Long, lastCol As Long
    Dim headers() As String, rows As New Collection, d As clsDict
    Dim wasUpdating As Boolean, omitNames As Boolean

    ' Every field forced to Text on the way in, mirroring modExport's "@"
    ' number format on the way out - so "00001" keeps its leading zeros and an
    ' ISO date/time string is not silently reparsed through the machine's
    ' locale before ParseIsoDateTime ever sees it.
    ReDim fi(1 To 60)
    For i = 1 To 60
        fi(i) = Array(i, xlTextFormat)
    Next i

    wasUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Workbooks.OpenText Filename:=path, Origin:=65001, StartRow:=1, DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False, _
        FieldInfo:=fi, TrailingMinusNumbers:=True
    Set wbIn = ActiveWorkbook
    Set ws = wbIn.Worksheets(1)

    lastCol = ws.Cells(HDR_ROW, ws.Columns.Count).End(xlToLeft).Column
    If Len(Trim$(CStr(ws.Cells(HDR_ROW, 1).Value))) = 0 Then
        wbIn.Close False
        Application.ScreenUpdating = wasUpdating
        ThisWorkbook.Activate
        Err.Raise vbObjectError + 515, "ReadImportRows", _
            "This file does not look like a print job export - no header row was found where one was expected."
    End If

    ' Marker written beside "Rows" by an export made without student names.
    omitNames = (StrComp(Trim$(CStr(ws.Cells(HDR_ROW - 1, 3).Value)), "Student names", vbTextCompare) = 0 And _
                 StrComp(Trim$(CStr(ws.Cells(HDR_ROW - 1, 4).Value)), "Omitted", vbTextCompare) = 0)

    ReDim headers(1 To lastCol)
    For c = 1 To lastCol
        headers(c) = Trim$(CStr(ws.Cells(HDR_ROW, c).Value))
    Next c

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then Exit For
        Set d = New clsDict
        For c = 1 To lastCol
            ' No "Student Name" key at all when the file left names out, so
            ' WriteImportedRow leaves the sheet's own names untouched.
            If Not (omitNames And StrComp(headers(c), "Student Name", vbTextCompare) = 0) Then
                d.Add headers(c), CStr(ws.Cells(r, c).Value)
            End If
        Next c
        rows.Add d
    Next r

    wbIn.Close False
    Application.ScreenUpdating = wasUpdating
    ThisWorkbook.Activate

    Set ReadImportRows = rows
End Function

' Manual parse, not CDate. CDate on "2026-09-15 10:12:00" is not itself
' locale-ambiguous (ISO order is unambiguous), but modExport writes this
' format specifically so no importer has to trust a date parser's guess - see
' modExport.CellOut's comment. Built the same defensive way as everywhere else
' this project turns a string into a number: only ever from digits it placed
' there itself.
Public Function ParseIsoDateTime(ByVal s As String) As Date
    s = Trim$(s)
    If Len(s) < 19 Then Exit Function
    Dim y As Integer, mo As Integer, da As Integer, hh As Integer, mi As Integer, se As Integer
    y = Val(Mid$(s, 1, 4))
    mo = Val(Mid$(s, 6, 2))
    da = Val(Mid$(s, 9, 2))
    hh = Val(Mid$(s, 12, 2))
    mi = Val(Mid$(s, 15, 2))
    se = Val(Mid$(s, 18, 2))
    ParseIsoDateTime = DateSerial(y, mo, da) + TimeSerial(hh, mi, se)
End Function
