Attribute VB_Name = "modBackup"
Option Explicit

' Full-workbook backup and restore (snag list item 4a) - the ad-hoc "get me
' back to where I was" path, distinct from modExport's per-location job CSVs
' (which already exist and are reused here unchanged) and Export report's
' point-in-time archive (§8.3/§10.4).
'
' Backup writes one CSV per catalogue table - Technicians, Printers, Papers,
' the four Settings-page lookup tables, and Settings itself - alongside the
' job CSVs Export All Locations already produces, all sharing one timestamp.
' Restore reads any ONE of those files back, finds every sibling sharing the
' same timestamp in the same folder, and reads the whole set back in:
' overwrite-by-stable-ID for the catalogue tables (the same "existing key:
' overwrite, new key: append" rule modImport already uses for job rows by
' Job ID), and the EXISTING modImport machinery, unmodified, for the job
' records themselves.
'
' Deliberately shaped to reuse modImport.ReadImportRows rather than writing a
' second CSV reader: a catalogue CSV's header block is padded to the same
' eight rows modExport.BuildBlock uses for a job export, so the table header
' always lands on row 9 and the data on row 10 - exactly where
' ReadImportRows' own hardcoded HDR_ROW/FIRST_DATA_ROW already look.

Private Const CSV_UTF8 As Long = 62         ' xlCSVUTF8

' The eight catalogue/configuration tables this backs up. Settings is one of
' them, not a separate mechanism - modUtils.Tbl finds it on the Settings
' sheet exactly like the other three lookup tables there.
Private Function CatalogTableNames() As Variant
    CatalogTableNames = Array("tblTechnicians", "tblPrinters", "tblPapers", _
        "tblPaperTypes", "tblStandardSizes", "tblPaperFamilies", "tblConsumables", "tblSettings")
End Function

' The column that identifies "the same row" across a backup and the live
' table, for overwrite-vs-append matching - the same role Job ID plays for
' modImport. tblPaperTypes/tblStandardSizes/tblPaperFamilies/tblConsumables
' have no synthetic ID column at all (§3.3); their natural-key text column is
' already what every lookup in this workbook treats as their identity
' (modCatalog's own dictionaries are keyed the same way), so reusing it here
' assumes nothing new.
Private Function CatalogKeyHeader(ByVal TableName As String) As String
    Select Case TableName
        Case "tblTechnicians":   CatalogKeyHeader = "TechID"
        Case "tblPrinters":      CatalogKeyHeader = "PrinterID"
        Case "tblPapers":        CatalogKeyHeader = "StockID"
        Case "tblPaperTypes":    CatalogKeyHeader = "Paper type"
        Case "tblStandardSizes": CatalogKeyHeader = "Size name"
        Case "tblPaperFamilies": CatalogKeyHeader = "Family"
        Case "tblConsumables":   CatalogKeyHeader = "Consumable type"
        Case "tblSettings":      CatalogKeyHeader = "Key"
    End Select
End Function

' ================================================================ backup ===
Public Sub BackupAll()
    Dim folder As String, ts As String, names As Variant, i As Long, lo As ListObject
    Dim done As Long, failed As String, path As String, why As String

    On Error GoTo Fail
    folder = ExportFolder()
    If Len(folder) = 0 Then
        Say "This workbook could not be backed up.", _
            "The workbook's own folder could not be resolved to a location on this computer. " & _
            "That happens when the file is open from OneDrive and Excel reports its address as a web link rather than a folder.", _
            "Open the workbook from the OneDrive folder on this computer rather than from the browser, then try again."
        Exit Sub
    End If

    ' Job records first, using the existing per-location export unchanged -
    ' its own summary dialog covers those. Catalogue tables get a second,
    ' separate summary below rather than trying to suppress and re-merge
    ' ExportAllLocations' own Say() call.
    ExportAllLocations

    ts = Format$(Now, "yyyymmdd-hhnn")
    names = CatalogTableNames()

    AppOff
    For i = LBound(names) To UBound(names)
        Set lo = Tbl(CStr(names(i)))
        If Not lo Is Nothing Then
            path = folder & Application.PathSeparator & "PrintCosts-" & SettingText("SITE_ID", "SITE") & _
                "-CATALOG-" & CStr(names(i)) & "-" & ts & ".csv"
            If WriteTableCsv(lo, path) Then
                done = done + 1
            Else
                failed = failed & "- " & names(i) & vbCrLf
            End If
        End If
    Next i
    AppOn

    why = done & " configuration table" & IIf(done = 1, "", "s") & " backed up (Technicians, Printers, Papers, the four Settings-page lookup tables, and Settings itself)."
    If Len(failed) > 0 Then why = why & vbCrLf & vbCrLf & "Could not be backed up:" & vbCrLf & failed

    Say "Catalogue backup complete.", why, _
        "Written to:" & vbCrLf & folder & vbCrLf & vbCrLf & _
        "Together with the location job exports above, Restore workbook can read this whole backup back in - point it at any one of this backup's files."
    Exit Sub
Fail:
    AppReset
    ReportError "Backup workbook"
End Sub

' Header block padded to the same 8 rows as modExport.BuildBlock's
' per-location export, so the table header lands on row 9 and data on row
' 10 - exactly what modImport.ReadImportRows' own hardcoded row numbers
' expect, letting Restore reuse that function unchanged for catalogue rows
' too, not just job rows.
Private Function WriteTableCsv(ByVal lo As ListObject, ByVal path As String) As Boolean
    Const HEAD As Long = 8
    Dim n As Long, w As Long, r As Long, c As Long, a() As Variant

    n = RowCount(lo)
    w = lo.ListColumns.Count
    ReDim a(1 To HEAD + 1 + n, 1 To w)

    a(1, 1) = "Print Cost Management catalogue backup"
    a(2, 1) = "Table":          a(2, 2) = lo.Name
    a(3, 1) = "Schema version": a(3, 2) = SCHEMA_VER
    a(4, 1) = "Site ID":        a(4, 2) = SettingText("SITE_ID", "SITE")
    a(5, 1) = "Site name":      a(5, 2) = SettingText("SITE_NAME")
    a(6, 1) = "Generated":      a(6, 2) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    a(7, 1) = "Rows":           a(7, 2) = n
    ' Row 8 deliberately blank - see the module comment on why HEAD is 8.

    For c = 1 To w
        a(HEAD + 1, c) = lo.ListColumns(c).Name
    Next c

    For r = 1 To n
        For c = 1 To w
            a(HEAD + 1 + r, c) = CellText(lo.DataBodyRange.Cells(r, c))
        Next c
    Next r

    WriteTableCsv = WriteBlockCsv(a, path)
End Function

' A genuinely blank cell's .Value2 is the number 0, not "" (the same
' phantom-zero trap modReports' Paid column hit, §8.3) - IsEmpty is checked
' first so a blank catalogue field (an optional Notes column, say) backs up
' as blank rather than the literal text "0".
Private Function CellText(ByVal c As Range) As Variant
    If IsEmpty(c.Value) Then
        CellText = ""
    Else
        CellText = c.Value2
    End If
End Function

' The actual staging-workbook-to-CSV write, shared by every table this backs
' up - same "force Text, write values, SaveAs xlCSVUTF8, confirm the file is
' actually on disk" technique as modExport's own exports (modExport.ExportOne/
' ExportReportSnapshot), not duplicated by copy-paste since this is the only
' other place in the workbook that writes a CSV this way.
Private Function WriteBlockCsv(ByRef block As Variant, ByVal path As String) As Boolean
    Dim wbOut As Workbook
    On Error GoTo Failed
    Set wbOut = Application.Workbooks.Add
    With wbOut.Worksheets(1)
        .Cells.NumberFormat = "@"
        .Range(.Cells(1, 1), .Cells(UBound(block, 1), UBound(block, 2))).Value = block
    End With
    wbOut.SaveAs path, CSV_UTF8
    wbOut.Close False
    WriteBlockCsv = FileExistsB(path)
    Exit Function
Failed:
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close False
    WriteBlockCsv = False
End Function

Private Function FileExistsB(ByVal p As String) As Boolean
    Dim a As Long
    On Error GoTo No
    a = GetAttr(p)
    FileExistsB = ((a And vbDirectory) <> vbDirectory)
    Exit Function
No:
    FileExistsB = False
End Function

' =============================================================== restore ===
' Entry point: the user picks any ONE file from a backup (GetOpenFilename,
' not Application.FileDialog - the same Mac-safe file picker modImport's own
' PickImportFile already uses, §9.3), and this finds every sibling sharing
' the same trailing "-yyyymmdd-hhmm.csv" in the same folder, classifies each
' one by its filename (this module wrote every filename it will ever read,
' so parsing them back is safe here in a way it would not be for an
' arbitrary file), and previews row counts before asking to confirm.
Public Sub RestoreWorkbook()
    Dim path As String
    On Error GoTo Fail
    path = PickBackupFile()
    If Len(path) = 0 Then Exit Sub
    RestoreFromFile path
    Exit Sub
Fail:
    AppReset
    ReportError "Restore workbook"
End Sub

' Split from RestoreWorkbook so a test can drive the real classify-preview-
' confirm path against a known file without the OS picker PickBackupFile
' opens (GetOpenFilename cannot run unattended, same reason modImport's own
' PickImportFile is never called directly by a test either).
Public Sub RestoreFromFile(ByVal AnyFilePath As String)
    Dim ts As String, folder As String
    Dim catalogPaths As clsDict, locationPaths As Collection
    Dim preview As String, k As Variant, n As Long, f As Variant, locCode As String

    ts = ExtractTimestamp(AnyFilePath)
    folder = LeftFolder(AnyFilePath)
    If Len(ts) = 0 Then
        Say "That does not look like a Print Cost Management backup file.", AnyFilePath, _
            "Choose a file named like PrintCosts-<site>-...-yyyymmdd-hhmm.csv, as Backup workbook writes them."
        Exit Sub
    End If

    ClassifyBackupSet folder, ts, catalogPaths, locationPaths

    If catalogPaths.Count = 0 And locationPaths.Count = 0 Then
        Say "No backup files were found alongside that one.", "Looked in:" & vbCrLf & folder
        Exit Sub
    End If

    For Each k In catalogPaths.Keys
        n = ReadImportRows(CStr(catalogPaths.Item(CStr(k)))).Count
        preview = preview & "- " & CStr(k) & ": " & n & " row" & IIf(n = 1, "", "s") & vbCrLf
    Next k
    For Each f In locationPaths
        locCode = PeekCell(CStr(f), 5, 2)
        n = ReadImportRows(CStr(f)).Count
        preview = preview & "- " & IIf(Len(locCode) > 0, locCode, "(unknown location)") & ": " & n & " job record" & IIf(n = 1, "", "s") & vbCrLf
    Next f

    If Not Ask("Restore the workbook from the backup taken " & FriendlyTimestamp(ts) & "?" & vbCrLf & vbCrLf & _
        preview & vbCrLf & _
        "Matching catalogue rows and settings will be OVERWRITTEN with the backup's version; rows not in the backup are left alone. " & _
        "Job records will be imported into each location listed above, following the same rule as Import (existing Job ID: overwritten; new Job ID: added). " & _
        "This cannot be undone.", "Restore workbook") Then Exit Sub

    RestoreWorkbookConfirmed catalogPaths, locationPaths
End Sub

' The actual write, given an already-classified file set. Public so a test
' can bypass BOTH the OS file picker and the Ask() confirm gate and drive
' the write directly - same reason modImport.ApplyImportConfirmed and
' modReports.DeleteVisibleReportsConfirmed are split this way. See
' RestoreFromFileConfirmed below for the version a test actually calls.
Private Sub ClassifyBackupSet(ByVal folder As String, ByVal ts As String, _
                              ByRef catalogPaths As clsDict, ByRef locationPaths As Collection)
    Dim files As Collection, f As Variant, fname As String, tableName As String
    Set catalogPaths = New clsDict
    Set locationPaths = New Collection

    Set files = MatchingBackupFiles(folder, ts)
    For Each f In files
        fname = FileNameOnly(CStr(f))
        tableName = CatalogFileTable(fname)
        If Len(tableName) > 0 Then
            catalogPaths.Add tableName, CStr(f)
        Else
            locationPaths.Add CStr(f)
        End If
    Next f
End Sub

' Classify-then-write with no preview and no confirm - what a test calls to
' exercise the real restore-application code (RestoreWorkbookConfirmed)
' without needing to construct a clsDict/Collection from outside VBA, which
' COM automation cannot do for this project's own class modules.
Public Sub RestoreFromFileConfirmed(ByVal AnyFilePath As String)
    Dim ts As String, folder As String
    Dim catalogPaths As clsDict, locationPaths As Collection

    ts = ExtractTimestamp(AnyFilePath)
    folder = LeftFolder(AnyFilePath)
    If Len(ts) = 0 Then Exit Sub

    ClassifyBackupSet folder, ts, catalogPaths, locationPaths
    RestoreWorkbookConfirmed catalogPaths, locationPaths
End Sub

' The actual write. Public so a test can call it directly, bypassing the
' Ask() gate above - Ask() always declines under SetQuiet, the same reason
' every other destructive command in this workbook (modJobs.RemoveRow,
' modImport.ApplyImportConfirmed, modReports.DeleteVisibleReportsConfirmed)
' is split into a confirm half and a do-it half.
Public Sub RestoreWorkbookConfirmed(ByVal catalogPaths As clsDict, ByVal locationPaths As Collection)
    Dim k As Variant, tableName As String, path As String, lo As ListObject
    Dim rows As Collection, appended As Long, overwrite As Long, skipped As Long
    Dim summary As String, f As Variant, ws As Worksheet, locCode As String

    On Error GoTo Fail
    AppOff

    For Each k In catalogPaths.Keys
        tableName = CStr(k)
        path = CStr(catalogPaths.Item(tableName))
        Set lo = Tbl(tableName)
        If Not lo Is Nothing Then
            Set rows = ReadImportRows(path)
            appended = 0: overwrite = 0: skipped = 0
            UnlockSheet lo.Parent
            ApplyCatalogRows tableName, lo, rows, appended, overwrite, skipped
            RelockSheet lo.Parent
            summary = summary & "- " & tableName & ": " & appended & " added, " & overwrite & " overwritten" & _
                IIf(skipped > 0, ", " & skipped & " skipped (read-only)", "") & vbCrLf
        End If
    Next k

    For Each f In locationPaths
        locCode = PeekCell(CStr(f), 5, 2)
        Set ws = SheetForCode(locCode)
        If Not ws Is Nothing Then
            Set rows = ReadImportRows(CStr(f))
            ApplyImportConfirmed ws, rows
            summary = summary & "- " & locCode & " (" & LocValue(ws, "LOC_Name") & "): " & rows.Count & " job record" & IIf(rows.Count = 1, "", "s") & " imported" & vbCrLf
        Else
            summary = summary & "- '" & locCode & "' - no matching print room sheet found, skipped." & vbCrLf
        End If
    Next f

    AppOn

    LogAudit "Restore workbook", "(multiple)", catalogPaths.Count & " catalogue table(s), " & locationPaths.Count & " location file(s)"

    Say "Workbook restore complete.", summary, _
        "Run Check workbook afterward to confirm everything looks right."
    Exit Sub
Fail:
    AppReset
    ReportError "Restore workbook"
End Sub

' Writes one catalogue/Settings row back by its stable key (CatalogKeyHeader):
' an existing key is overwritten, a new one appended - identical shape to
' modImport's Job ID rule, just against a different key column per table.
'
' Settings rows marked "Read-only" in their own Notes column (modInit.
' UnlockSettingsValues' own convention - APP_VER, SCHEMA, BUILT, BUILT_BY,
' LASTREF) are skipped outright: those describe the CURRENT build, and a
' restore rewinding them to a backup's old values would make the workbook
' misreport its own version and schema.
Private Sub ApplyCatalogRows(ByVal TableName As String, ByVal lo As ListObject, ByVal rows As Collection, _
                             ByRef Appended As Long, ByRef Overwrite As Long, ByRef Skipped As Long)
    Dim keyHdr As String, d As clsDict, key As String, n As Long, v As Variant
    Dim notesVal As String, skip As Boolean

    keyHdr = CatalogKeyHeader(TableName)
    If Len(keyHdr) = 0 Then Exit Sub

    For Each v In rows
        Set d = v
        key = Trim$(CStr(d.Item(keyHdr)))
        If Len(key) > 0 Then
            skip = False
            If StrComp(TableName, "tblSettings", vbTextCompare) = 0 Then
                If d.Exists("Notes") Then notesVal = Trim$(CStr(d.Item("Notes"))) Else notesVal = ""
                skip = (Left$(notesVal, 9) = "Read-only")
            End If

            If skip Then
                Skipped = Skipped + 1
            Else
                n = FindRowByKey(lo, keyHdr, key)
                If n = 0 Then
                    Appended = Appended + 1
                    If lo.ListRows.Count = 1 And IsBlankRow(lo, 1) Then
                        n = 1
                    Else
                        n = lo.ListRows.Add.Index
                    End If
                Else
                    Overwrite = Overwrite + 1
                End If
                WriteCatalogRow lo, n, d
            End If
        End If
    Next v
End Sub

Private Function FindRowByKey(ByVal lo As ListObject, ByVal KeyHeader As String, ByVal Key As String) As Long
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, KeyHeader).Value)), Key, vbTextCompare) = 0 Then
            FindRowByKey = i
            Exit Function
        End If
    Next i
End Function

' Every column present in the backup row is written except one currently
' holding a FORMULA (tblPapers' Measure/Cost unit, §3.3) - checked by
' .HasFormula rather than a hardcoded per-table column list, so a future
' calculated column added to any catalogue table is protected automatically.
' A value that parses as a number is written as one (CDbl, locale-aware -
' modUtils.NumOf's own comment on why Val() is wrong here applies equally),
' so cost/width/height columns stay real numbers for downstream formulas
' rather than becoming text that happens to look like one.
Private Sub WriteCatalogRow(ByVal lo As ListObject, ByVal RowNo As Long, ByVal d As clsDict)
    Dim c As Long, header As String, s As String, cell As Range
    For c = 1 To lo.ListColumns.Count
        header = lo.ListColumns(c).Name
        Set cell = CellIn(lo, RowNo, header)
        If Not cell.HasFormula Then
            s = ""
            If d.Exists(header) Then s = Trim$(CStr(d.Item(header)))
            If Len(s) = 0 Then
                cell.ClearContents
            ElseIf IsNumeric(s) Then
                cell.Value = CDbl(s)
            Else
                cell.Value = s
            End If
        End If
    Next c
End Sub

' ============================================================== file set ===
Private Function PickBackupFile() As String
    Dim f As Variant
    f = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Restore workbook - choose any one file from the backup")
    If VarType(f) = vbBoolean Then Exit Function   ' Cancel
    PickBackupFile = CStr(f)
End Function

Private Function MatchingBackupFiles(ByVal folder As String, ByVal ts As String) As Collection
    Dim out As New Collection, f As String
    f = Dir$(folder & Application.PathSeparator & "PrintCosts-*-" & ts & ".csv")
    Do While Len(f) > 0
        out.Add folder & Application.PathSeparator & f
        f = Dir$()
    Loop
    Set MatchingBackupFiles = out
End Function

' Filenames are parsed here, safely, only because this module wrote every
' filename it will ever read back (BackupAll's own naming, above) - not a
' general-purpose filename parser, and not used for anything modExport's own
' per-location export writes (those are told apart by absence of "-CATALOG-",
' not by parsing their own name further).
Private Function CatalogFileTable(ByVal fname As String) As String
    Dim names As Variant, i As Long
    names = CatalogTableNames()
    For i = LBound(names) To UBound(names)
        If InStr(1, fname, "-CATALOG-" & CStr(names(i)) & "-", vbTextCompare) > 0 Then
            CatalogFileTable = CStr(names(i))
            Exit Function
        End If
    Next i
End Function

Private Function FileNameOnly(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, Application.PathSeparator)
    If i > 0 Then FileNameOnly = Mid$(p, i + 1) Else FileNameOnly = p
End Function

Private Function LeftFolder(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, Application.PathSeparator)
    If i > 0 Then LeftFolder = Left$(p, i - 1)
End Function

' The last two "-"-delimited tokens before ".csv" are always the date and
' time fields (Format$(Now,"yyyymmdd-hhnn")) in every filename this module or
' modExport writes - safe to assume no other component (site ID, table name,
' location code) ever contains a hyphen, since none of this workbook's own
' generated codes do.
Private Function ExtractTimestamp(ByVal path As String) As String
    Dim fname As String, noExt As String, parts() As String, n As Long
    Dim datePart As String, timePart As String

    fname = FileNameOnly(path)
    If Len(fname) < 5 Or LCase$(Right$(fname, 4)) <> ".csv" Then Exit Function
    If Left$(fname, 11) <> "PrintCosts-" Then Exit Function

    noExt = Left$(fname, Len(fname) - 4)
    parts = Split(noExt, "-")
    n = UBound(parts)
    If n < 2 Then Exit Function

    datePart = parts(n - 1)
    timePart = parts(n)
    If Len(datePart) <> 8 Or Not IsNumeric(datePart) Then Exit Function
    If Len(timePart) <> 4 Or Not IsNumeric(timePart) Then Exit Function

    ExtractTimestamp = datePart & "-" & timePart
End Function

Private Function FriendlyTimestamp(ByVal ts As String) As String
    Dim d As String, t As String
    d = Left$(ts, 8)
    t = Mid$(ts, 10, 4)
    FriendlyTimestamp = Mid$(d, 7, 2) & "/" & Mid$(d, 5, 2) & "/" & Left$(d, 4) & " " & Left$(t, 2) & ":" & Right$(t, 2)
End Function

' Opens a CSV just far enough to read one cell (the "Location code" line, row
' 5, in a job export's own header block, modExport.BuildBlock) - not routed
' through ReadImportRows, which only returns the DATA rows and has no way to
' hand back a header-block value.
Private Function PeekCell(ByVal path As String, ByVal r As Long, ByVal c As Long) As String
    Dim wbIn As Workbook, ws As Worksheet, v As String, wasUpdating As Boolean
    On Error GoTo Failed
    wasUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Workbooks.OpenText Filename:=path, Origin:=65001, StartRow:=1, DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False
    Set wbIn = ActiveWorkbook
    Set ws = wbIn.Worksheets(1)
    v = Trim$(CStr(ws.Cells(r, c).Value))
    wbIn.Close False
    Application.ScreenUpdating = wasUpdating
    ThisWorkbook.Activate
    PeekCell = v
    Exit Function
Failed:
    On Error Resume Next
    If Not wbIn Is Nothing Then wbIn.Close False
    Application.ScreenUpdating = wasUpdating
    ThisWorkbook.Activate
End Function
