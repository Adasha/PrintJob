Attribute VB_Name = "modBackup"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Full-workbook backup and restore (snag list item 4a) - the ad-hoc "get me
' back to where I was" path, distinct from modExport's per-location job CSVs
' (which already exist and are reused here unchanged) and Export report's
' point-in-time archive (§8.3/§10.4).
'
' Backup writes one CSV per catalogue table - Technicians, Printers, Papers,
' the four Settings-page lookup tables, and Settings itself - alongside the
' job CSVs Export All Locations already produces, all sharing one timestamp.
' Restore reads any ONE of those files back, finds every sibling sharing the
' same timestamp in the same folder, creates any print room the backup has
' and this workbook lacks, and reads the whole set back in:
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

' The shipped sample room - see "the example room" below.
Private Const EXAMPLE_ROOM_NAME As String = "Example Print Room"
Private Const EXAMPLE_ROOM_SAMPLE_JOBS As Long = 12

' The eight catalogue/configuration tables this backs up. Settings is one of
' them, not a separate mechanism - modUtils.Tbl finds it on the Settings
' sheet exactly like the other three lookup tables there. tblColours (0.11.0)
' and tblDepartments (0.10.30) are the newest; a backup taken before either
' simply has no file for it, and a restore skips any table the backup lacks.
Private Function CatalogTableNames() As Variant
    CatalogTableNames = Array("tblTechnicians", "tblPrinters", "tblPapers", "tblDepartments", "tblColours", _
        "tblPaperTypes", "tblStandardSizes", "tblConsumables", "tblSettings")
End Function

' The column that identifies "the same row" across a backup and the live
' table, for overwrite-vs-append matching - the same role Job ID plays for
' modImport. tblPaperTypes/tblStandardSizes/tblConsumables
' have no synthetic ID column at all (§3.3); their natural-key text column is
' already what every lookup in this workbook treats as their identity
' (modCatalog's own dictionaries are keyed the same way), so reusing it here
' assumes nothing new.
Private Function CatalogKeyHeader(ByVal TableName As String) As String
    Select Case TableName
        Case "tblTechnicians":   CatalogKeyHeader = "TechID"
        Case "tblPrinters":      CatalogKeyHeader = "PrinterID"
        Case "tblPapers":        CatalogKeyHeader = "StockID"
        Case "tblDepartments":   CatalogKeyHeader = "DeptID"
        Case "tblColours":       CatalogKeyHeader = "ColourID"
        Case "tblPaperTypes":    CatalogKeyHeader = "Paper type"
        Case "tblStandardSizes": CatalogKeyHeader = "Size name"
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
    ExportAllLocations False

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

    why = done & " configuration table" & IIf(done = 1, "", "s") & " backed up (Technicians, Printers, Papers, Departments, Consumables colours, the three Settings-page lookup tables, and Settings itself)."
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
    Set wbOut = NewScratchBook()   ' not Workbooks.Add - see modUtils.NewScratchBook
    With wbOut.Worksheets(1)
        .Cells.NumberFormat = "@"
        .Range(.Cells(1, 1), .Cells(UBound(block, 1), UBound(block, 2))).Value = block
    End With
    wbOut.SaveAs path, CSV_UTF8
    CloseScratchBook wbOut
    WriteBlockCsv = FileWritten(path)
    Exit Function
Failed:
    On Error Resume Next
    If Not wbOut Is Nothing Then CloseScratchBook wbOut
    WriteBlockCsv = False
End Function

' =============================================================== restore ===
' Entry point: the user picks any ONE file from a backup (GetOpenFilename,
' not Application.FileDialog - the same Mac-safe file picker modImport's own
' PickImportFile already uses, §9.3), and this finds every sibling sharing
' the same trailing "-yyyymmdd-hhmm.csv" in the same folder, classifies each
' one by its filename (this module wrote every filename it will ever read,
' so parsing them back is safe here in a way it would not be for an
' arbitrary file), and previews row counts before asking to confirm.
'
' The one file can also be an OLDER COPY OF THE WORKBOOK (.xlsm/.xlsx/.xlsb):
' the upgrade path is a fresh workbook plus a restore, and the old workbook
' already holds everything a restore needs - every catalogue table, each
' print room's settings and every room's job records - so nothing has to be
' backed up first.
'
' Whatever the source, it is first read into one "restore plan" - the
' catalogue rows per table, plus one clsRestoreRoom per print room - and the
' plan is then previewed and applied by the same code, so every source gets
' the same behaviour: a print room the backup mentions that does not exist
' here yet is CREATED (modRegistry.CreatePrintRoom) before its job records
' go in, rather than being skipped.
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
    Dim catalogRows As clsDict, rooms As Collection, srcDesc As String

    If Not LoadRestorePlan(AnyFilePath, catalogRows, rooms, srcDesc) Then Exit Sub

    If Not Ask("Restore the workbook from " & srcDesc & "?" & vbCrLf & vbCrLf & _
        PlanPreview(catalogRows, rooms) & vbCrLf & _
        "Matching catalogue rows and settings will be OVERWRITTEN with the backup's version; rows not in the backup are left alone. " & _
        "A print room listed above that does not exist here yet will be CREATED" & _
        IIf(PlanHasRoomSettings(rooms), ", with its default technician, printer and paper, permitted printers and roll length unit from the workbook. ", ". ") & _
        "Job records will be imported into each print room listed above, following the same rule as Import (existing Job ID: overwritten; new Job ID: added). " & _
        IIf(ExampleRoomToRemove(rooms) Is Nothing, "", "The sample '" & EXAMPLE_ROOM_NAME & "' print room, which this restore does not touch, will be REMOVED once the restore completes. ") & _
        "This cannot be undone.", "Restore workbook") Then Exit Sub

    ApplyRestorePlan catalogRows, rooms
End Sub

' Load-then-write with no preview and no confirm - what a test calls to
' exercise the real restore-application code (ApplyRestorePlan) without
' needing the Ask() gate, which always declines under SetQuiet, the same
' reason every other destructive command in this workbook (modJobs.RemoveRow,
' modImport.ApplyImportConfirmed, modReports.DeleteVisibleReportsConfirmed)
' is split into a confirm half and a do-it half.
Public Sub RestoreFromFileConfirmed(ByVal AnyFilePath As String)
    Dim catalogRows As clsDict, rooms As Collection, srcDesc As String

    If Not LoadRestorePlan(AnyFilePath, catalogRows, rooms, srcDesc) Then Exit Sub
    ApplyRestorePlan catalogRows, rooms
End Sub

' ============================================================ restore plan ===
' Reads the chosen file (and, for a CSV backup, its siblings) into memory:
'   CatalogRows  table name -> Collection of clsDict rows keyed by header
'   Rooms        Collection of clsRestoreRoom, one per print room
'   SrcDesc      how the preview dialog names the source
' False (after telling the user why) when there is nothing usable to restore.
Private Function LoadRestorePlan(ByVal AnyFilePath As String, ByRef CatalogRows As clsDict, _
                                 ByRef Rooms As Collection, ByRef SrcDesc As String) As Boolean
    Set CatalogRows = New clsDict
    Set Rooms = New Collection
    If IsWorkbookFile(AnyFilePath) Then
        LoadRestorePlan = LoadFromWorkbook(AnyFilePath, CatalogRows, Rooms, SrcDesc)
    Else
        LoadRestorePlan = LoadFromCsvSet(AnyFilePath, CatalogRows, Rooms, SrcDesc)
    End If
End Function

Private Function LoadFromCsvSet(ByVal AnyFilePath As String, ByVal CatalogRows As clsDict, _
                                ByVal Rooms As Collection, ByRef SrcDesc As String) As Boolean
    Dim ts As String, folder As String
    Dim catalogPaths As clsDict, locationPaths As Collection
    Dim k As Variant, f As Variant, room As clsRestoreRoom

    ts = ExtractTimestamp(AnyFilePath)
    folder = LeftFolder(AnyFilePath)
    If Len(ts) = 0 Then
        Say "That does not look like a Print Cost Management backup file.", AnyFilePath, _
            "Choose a file named like PrintCosts-<site>-...-yyyymmdd-hhmm.csv, as Backup workbook writes them."
        Exit Function
    End If

    ClassifyBackupSet folder, ts, catalogPaths, locationPaths

    If catalogPaths.Count = 0 And locationPaths.Count = 0 Then
        Say "No backup files were found alongside that one.", "Looked in:" & vbCrLf & folder
        Exit Function
    End If

    For Each k In catalogPaths.Keys
        CatalogRows.Add CStr(k), ReadImportRows(CStr(catalogPaths.Item(CStr(k))))
    Next k

    ' A job export's header block names its room (rows 5 and 6) and, from
    ' 0.10.23, its department beside the name (row 6, column 4) - enough to
    ' create the room if it is missing. Older files simply have no department.
    For Each f In locationPaths
        Set room = New clsRestoreRoom
        room.Code = PeekCell(CStr(f), 5, 2)
        room.RoomName = PeekCell(CStr(f), 6, 2)
        room.Dept = PeekCell(CStr(f), 6, 4)
        Set room.JobRows = ReadImportRows(CStr(f))
        Rooms.Add room
    Next f

    SrcDesc = "the backup taken " & FriendlyTimestamp(ts)
    LoadFromCsvSet = True
End Function

' ---------------------------------------------------------- older workbook ---
' Reads an older copy of the workbook straight into the same plan a CSV
' backup gives: the seven catalogue tables, then for each print room its
' LOC_* settings and job records. Nothing is written to the older file - it is
' opened read-only, from a temporary copy, with its macros off.
Private Function LoadFromWorkbook(ByVal Path As String, ByVal CatalogRows As clsDict, _
                                  ByVal Rooms As Collection, ByRef SrcDesc As String) As Boolean
    Dim wb As Workbook, tmp As String, ws As Worksheet, lo As ListObject
    Dim names As Variant, locs As Variant, i As Long, room As clsRestoreRoom
    Dim errNo As Long, errMsg As String

    If Not FileExists(Path) Then
        Say "That file could not be found.", Path
        Exit Function
    End If
    If StrComp(Path, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Say "That is this workbook.", Path, "Choose an older copy of the workbook, or one of the CSV files from a backup."
        Exit Function
    End If

    On Error GoTo Fail
    Set wb = OpenSourceWorkbook(Path, tmp)

    names = CatalogTableNames()
    For i = LBound(names) To UBound(names)
        Set lo = TableIn(wb, CStr(names(i)))
        If Not lo Is Nothing Then
            CatalogRows.Add CStr(names(i)), TableRows(lo, CatalogKeyHeader(CStr(names(i))))
        End If
    Next i

    locs = RoomSettingNames()
    For Each ws In wb.Worksheets
        If IsLocation(ws) Then
            Set lo = JobsTable(ws)
            If Not lo Is Nothing Then
                Set room = New clsRestoreRoom
                room.Code = Trim$(LocValue(ws, "LOC_Code"))
                room.RoomName = Trim$(LocValue(ws, "LOC_Name"))
                room.Dept = Trim$(LocValue(ws, "LOC_Dept"))
                Set room.Settings = New clsDict
                For i = LBound(locs) To UBound(locs)
                    room.Settings.Add CStr(locs(i)), LocValue(ws, CStr(locs(i)))
                Next i
                Set room.JobRows = TableRows(lo, "Job ID")
                Rooms.Add room
            End If
        End If
    Next ws

    CloseSource wb, tmp
    On Error GoTo 0

    If CatalogRows.Count = 0 And Rooms.Count = 0 Then
        Say "That workbook does not look like a Print Cost Management workbook.", Path, _
            "Choose a copy of this workbook from an earlier version, or one of the CSV files from a backup."
        Exit Function
    End If

    SrcDesc = "the workbook '" & FileNameOnly(Path) & "'"
    LoadFromWorkbook = True
    Exit Function
Fail:
    errNo = Err.Number: errMsg = Err.Description
    CloseSource wb, tmp
    Err.Raise errNo, "LoadFromWorkbook", errMsg
End Function

' The per-room settings that live on a room's sheet rather than in a table.
' LOC_Code, LOC_Name and LOC_Dept are read separately (they identify the room);
' these are what a new room is given beyond that.
Private Function RoomSettingNames() As Variant
    RoomSettingNames = Array("LOC_DefTech", "LOC_DefPrinter", "LOC_DefPaper", _
        "LOC_DefDisPaper", "LOC_DefDisCons", "LOC_Printers", "LOC_RollUnit")
End Function

Private Function PlanHasRoomSettings(ByVal Rooms As Collection) As Boolean
    Dim v As Variant, room As clsRestoreRoom
    For Each v In Rooms
        Set room = v
        If Not room.Settings Is Nothing Then
            PlanHasRoomSettings = True
            Exit Function
        End If
    Next v
End Function

Private Function IsWorkbookFile(ByVal Path As String) As Boolean
    Select Case LCase$(Mid$(Path, InStrRev(Path, ".") + 1))
        Case "xlsm", "xlsx", "xlsb"
            IsWorkbookFile = True
    End Select
End Function

' Opens the older workbook for reading only, without anything in it running.
'   - From a temporary COPY. Excel will not open two workbooks with the same
'     file name, and the older copy usually has the new one's name (PrintJob.xlsm,
'     in another folder); a copy also copes with the older one being open or
'     syncing. If the copy cannot be made the original is opened read-only -
'     unless that name clashes, which is then reported.
'   - Macros disabled (AutomationSecurity) and events off, so the older
'     workbook's own Workbook_Open does not run: it would re-protect, re-bind
'     and recalculate a workbook that is only being read.
' The caller closes it with CloseSource.
Private Function OpenSourceWorkbook(ByVal Path As String, ByRef TempPath As String) As Workbook
    Dim openPath As String, prevSec As Long, prevEvents As Boolean, wb As Workbook
    Dim errNo As Long, errMsg As String

    TempPath = CopyToTemp(Path)
    If Len(TempPath) > 0 Then
        openPath = TempPath
    Else
        If StrComp(FileNameOnly(Path), ThisWorkbook.Name, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 516, "OpenSourceWorkbook", _
                "The older workbook has the same file name as this one ('" & ThisWorkbook.Name & "') and a temporary copy of it could not be made, " & _
                "so Excel cannot open both. Rename the older file (for example add 'old' to its name) and try again."
        End If
        openPath = Path
    End If

    ' Mac Excel: whether AutomationSecurity is honoured there is unverified,
    ' so the older file's macro prompt may still appear. Say so up front
    ' rather than leaving the user wondering (0.10.25 test note 3). Through
    ' Say, so quiet mode (a test run) logs it instead of stopping on it.
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        Say "Excel may now ask whether to enable macros for the older workbook.", "", _
            "Choose Disable Macros: they are not needed, and the data is read without them.", _
            "Importing from an older workbook"
    End If

    prevEvents = Application.EnableEvents
    On Error Resume Next
    prevSec = Application.AutomationSecurity
    Application.AutomationSecurity = 3          ' msoAutomationSecurityForceDisable
    On Error GoTo Failed
    Application.EnableEvents = False
    Set wb = Workbooks.Open(Filename:=openPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Application.EnableEvents = prevEvents
    On Error Resume Next
    Application.AutomationSecurity = prevSec
    On Error GoTo 0
    ThisWorkbook.Activate
    Set OpenSourceWorkbook = wb
    Exit Function
Failed:
    errNo = Err.Number: errMsg = Err.Description
    Application.EnableEvents = prevEvents
    On Error Resume Next
    Application.AutomationSecurity = prevSec
    If Len(TempPath) > 0 Then Kill TempPath
    TempPath = ""
    On Error GoTo 0
    Err.Raise errNo, "OpenSourceWorkbook", errMsg
End Function

Private Sub CloseSource(ByRef Wb As Workbook, ByVal TempPath As String)
    On Error Resume Next
    If Not Wb Is Nothing Then Wb.Close False
    Set Wb = Nothing
    If Len(TempPath) > 0 Then Kill TempPath
    ThisWorkbook.Activate
    On Error GoTo 0
End Sub

' A uniquely named copy in the temp folder, or "" when there is no usable temp
' folder or the copy fails (a file Excel has locked, a read-only temp folder).
Private Function CopyToTemp(ByVal Src As String) As String
    Dim tmpDir As String, dst As String, ext As String
    tmpDir = Environ$("TEMP")
    If Len(tmpDir) = 0 Then tmpDir = Environ$("TMPDIR")
    If Len(tmpDir) = 0 Then tmpDir = Environ$("TMP")
    If Len(tmpDir) = 0 Then Exit Function
    If Right$(tmpDir, 1) = Application.PathSeparator Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)

    ext = Mid$(Src, InStrRev(Src, "."))
    dst = tmpDir & Application.PathSeparator & "PrintCosts-restore-" & Format$(Now, "yyyymmdd-hhnnss") & ext

    On Error GoTo Failed
    FileCopy Src, dst
    If FileExists(dst) Then CopyToTemp = dst
    Exit Function
Failed:
    CopyToTemp = ""
End Function

' modUtils.Tbl, but for another workbook.
Private Function TableIn(ByVal Wb As Workbook, ByVal TableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In Wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, TableName, vbTextCompare) = 0 Then
                Set TableIn = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' A table's data rows as the same clsDict-per-row (keyed by header text) that
' modImport.ReadImportRows gives for a CSV, so everything downstream is
' shared. Rows with no key are left out, as ReadImportRows stops at one. Read
' in a single array fetch: a job table can be thousands of rows by dozens of
' columns, and a COM call per cell would take minutes.
Private Function TableRows(ByVal lo As ListObject, ByVal KeyHeader As String) As Collection
    Dim out As Collection, d As clsDict, a As Variant, hdr() As String
    Dim r As Long, c As Long, keyCol As Long, nCols As Long, parentCol As Long, isData As Boolean

    Set out = New Collection
    Set TableRows = out
    If Not ColumnExists(lo, KeyHeader) Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    keyCol = ColIdx(lo, KeyHeader)
    ' A job table's colour-pass rows have no Job ID, only a Parent, so a row there
    ' counts as data when either is filled in (multi-pass: passes share the table).
    If StrComp(KeyHeader, "Job ID", vbTextCompare) = 0 And ColumnExists(lo, "Parent") Then parentCol = ColIdx(lo, "Parent")
    nCols = lo.ListColumns.Count
    ReDim hdr(1 To nCols)
    For c = 1 To nCols
        hdr(c) = lo.ListColumns(c).Name
    Next c

    a = lo.DataBodyRange.Value2
    If Not IsArray(a) Then Exit Function
    For r = 1 To UBound(a, 1)
        isData = (Len(Trim$(SourceText(a(r, keyCol), KeyHeader))) > 0)
        If Not isData And parentCol > 0 Then isData = (Len(Trim$(SourceText(a(r, parentCol), "Parent"))) > 0)
        If isData Then
            Set d = New clsDict
            For c = 1 To nCols
                d.Add hdr(c), SourceText(a(r, c), hdr(c))
            Next c
            out.Add d
        End If
    Next r
End Function

' One cell as the text a CSV backup would hold for it: blank for an empty or
' error cell, and the two timestamp columns as yyyy-mm-dd hh:nn:ss exactly as
' modExport.CellOut writes them (modImport.ParseIsoDateTime reads that back;
' a serial number or a locale-formatted date would not survive).
Private Function SourceText(ByVal V As Variant, ByVal Header As String) As String
    If IsError(V) Then Exit Function
    If IsEmpty(V) Then Exit Function
    If StrComp(Header, "Date/Time", vbTextCompare) = 0 Or StrComp(Header, "S_StampedAt", vbTextCompare) = 0 Then
        If VarType(V) = vbDouble Then
            If CDbl(V) > 0 Then SourceText = Format$(CDate(CDbl(V)), "yyyy-mm-dd hh:nn:ss")
            Exit Function
        End If
    End If
    SourceText = CStr(V)
End Function

Private Function PlanPreview(ByVal CatalogRows As clsDict, ByVal Rooms As Collection) As String
    Dim k As Variant, v As Variant, n As Long, room As clsRestoreRoom, s As String
    Dim tableRows As Collection, label As String

    For Each k In CatalogRows.Keys
        Set tableRows = CatalogRows.Obj(CStr(k))
        n = tableRows.Count
        s = s & "- " & CStr(k) & ": " & n & " row" & IIf(n = 1, "", "s") & vbCrLf
    Next k

    For Each v In Rooms
        Set room = v
        n = room.JobRows.Count
        If Len(room.Code) = 0 Then
            label = "(unknown location)"
        Else
            label = room.Code
        End If
        s = s & "- " & label & ": " & n & " job record" & IIf(n = 1, "", "s")
        If Len(room.Code) = 0 Then
            s = s & " (no location code - will be skipped)"
        ElseIf SheetForCode(room.Code) Is Nothing Then
            s = s & " (print room will be created)"
        End If
        s = s & vbCrLf
    Next v
    PlanPreview = s
End Function

' Private: it takes a clsDict/Collection built by the loaders above, which COM
' automation cannot construct for this project's own class modules - a test
' drives RestoreFromFileConfirmed instead.
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

' ================================================================== apply ===
' The actual write, in three passes so each depends only on the one before:
'   1. catalogue tables (so the printers a room's jobs name exist),
'   2. print rooms - looked up by code, created if missing,
'   3. job records, into each room, by the existing modImport machinery.
Private Sub ApplyRestorePlan(ByVal CatalogRows As clsDict, ByVal Rooms As Collection)
    Dim k As Variant, v As Variant, tableName As String, lo As ListObject
    Dim rows As Collection, appended As Long, overwrite As Long, skipped As Long, renamed As Long
    Dim summary As String, room As clsRestoreRoom, made As Long, ran As Long
    Dim back As Object

    On Error GoTo Fail
    Set back = ActiveSheet

    AppOff
    For Each k In CatalogRows.Keys
        tableName = CStr(k)
        Set lo = Tbl(tableName)
        If Not lo Is Nothing Then
            Set rows = CatalogRows.Obj(tableName)
            appended = 0: overwrite = 0: skipped = 0: renamed = 0
            UnlockSheet lo.Parent
            ApplyCatalogRows tableName, lo, rows, appended, overwrite, skipped, renamed
            RelockSheet lo.Parent
            summary = summary & "- " & tableName & ": " & appended & " added, " & overwrite & " overwritten" & _
                IIf(renamed > 0, ", " & renamed & " renamed to avoid a duplicate name", "") & _
                IIf(skipped > 0, ", " & skipped & " skipped (read-only)", "") & vbCrLf
        End If
    Next k
    AppOn

    ' Rooms are prepared outside the AppOff bracket above and below:
    ' CreatePrintRoom brackets its own work.
    For Each v In Rooms
        Set room = v
        PrepareRoom room, summary, made
    Next v

    AppOff
    For Each v In Rooms
        Set room = v
        If Not room.Sheet Is Nothing Then
            ApplyImportConfirmed room.Sheet, room.JobRows
            ran = ran + 1
            summary = summary & "- " & room.Code & " (" & LocValue(room.Sheet, "LOC_Name") & "): " & _
                IIf(room.Created, "print room created, ", "") & _
                room.JobRows.Count & " job record" & IIf(room.JobRows.Count = 1, "", "s") & " imported" & vbCrLf
        End If
    Next v
    AppOn

    ' Only when every room in the backup landed: a skipped room means the
    ' restore is incomplete, and the sample room is not worth losing for that.
    If ran > 0 And ran = Rooms.Count Then RemoveExampleRoom Rooms, back, summary

    ' CreatePrintRoom leaves the new room active; put the user back where they were.
    On Error Resume Next
    If Not back Is Nothing Then back.Activate
    On Error GoTo Fail

    LogAudit "Restore workbook", "(multiple)", CatalogRows.Count & " catalogue table(s), " & ran & " location(s)" & _
        IIf(made > 0, ", " & made & " print room(s) created", "")

    Say "Workbook restore complete.", summary, _
        "Run Check workbook afterward to confirm everything looks right."
    Exit Sub
Fail:
    AppReset
    ReportError "Restore workbook"
End Sub

' ---------------------------------------------------- the example room ---
' The workbook ships with one sample room (12 invented jobs) so the features can
' be tried before any real room exists. Once a restore has brought in the site's
' real rooms it is only clutter - and a room that would turn up in Reports and in
' the next Backup - so the restore removes it.
'
' The sample is recognised by what it is, not by its code: the shipped name, and
' nothing but sample jobs (a real job gets a Job ID after the sample ones). Its
' code is EXAMPLE now; workbooks built before 0.11.1 shipped it as MAIN, and a
' site's own first room is as likely to be coded MAIN as anything, so no code is
' assumed here.

' The sample room's sheet when this plan should remove it, else Nothing. Rooms
' the restore itself writes into are never candidates: a restored room that
' happens to be called Example Print Room is the site's, not the sample.
Private Function ExampleRoomToRemove(ByVal Rooms As Collection) As Worksheet
    Dim ws As Worksheet, v As Variant, rv As Variant, room As clsRestoreRoom, real As Long, restored As Boolean

    For Each v In Rooms
        Set room = v
        If Len(room.Code) > 0 Then real = real + 1
    Next v
    If real = 0 Then Exit Function

    For Each v In LocationSheets()
        Set ws = v
        If IsSampleRoom(ws) Then
            restored = False
            For Each rv In Rooms
                Set room = rv
                If StrComp(room.Code, LocValue(ws, "LOC_Code"), vbTextCompare) = 0 Then restored = True
            Next rv
            If Not restored Then
                Set ExampleRoomToRemove = ws
                Exit Function
            End If
        End If
    Next v
End Function

' The sample room's sheet when Room is the backup's own room under the sample's
' code and the sample is still untouched. Such a room must not be restored INTO
' the sample (it would keep the sample's name, defaults and permitted printers),
' so PrepareRoom reuses the sheet for it. That includes a backup whose room is
' itself still called Example Print Room: a site that kept the sample and added
' its own data to it (and its own permitted printers) - an unrenamed sample is
' the sample's sheet, not the user's.
Private Function SampleRoomToReuse(ByVal Room As clsRestoreRoom) As Worksheet
    Dim ws As Worksheet

    If Len(Room.RoomName) = 0 Then Exit Function
    Set ws = SheetForCode(Room.Code)
    If ws Is Nothing Then Exit Function
    If IsSampleRoom(ws) Then Set SampleRoomToReuse = ws
End Function

' Still the sample as shipped: its name, and no job but the sample ones.
Private Function IsSampleRoom(ByVal Ws As Worksheet) As Boolean
    If StrComp(LocValue(Ws, "LOC_Name"), EXAMPLE_ROOM_NAME, vbTextCompare) <> 0 Then Exit Function
    IsSampleRoom = HoldsOnlySampleJobs(Ws)
End Function

Private Function HoldsOnlySampleJobs(ByVal Ws As Worksheet) As Boolean
    Dim lo As ListObject, c As Range, id As String

    Set lo = JobsTable(Ws)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count > EXAMPLE_ROOM_SAMPLE_JOBS Then Exit Function
    If lo.ListRows.Count > 0 Then
        For Each c In lo.ListColumns("Job ID").DataBodyRange.Cells
            id = Trim$(CStr(c.Value2))
            If Len(id) > 0 Then
                If Not id Like "*-" & LocValue(Ws, "LOC_Code") & "-000##" Then Exit Function
                If Val(Right$(id, 2)) > EXAMPLE_ROOM_SAMPLE_JOBS Then Exit Function
            End If
        Next c
    End If
    HoldsOnlySampleJobs = True
End Function

' Deletes the example room through the same core Remove print room uses (audit
' entry and Refresh Locations included; the refresh dialog is muted so it does
' not bury the restore's own summary). Back is the sheet to return the user to;
' it is repointed at the first restored room if it was the one being deleted.
Private Sub RemoveExampleRoom(ByVal Rooms As Collection, ByRef Back As Object, ByRef Summary As String)
    Dim ex As Worksheet, wasQuiet As Boolean, room As clsRestoreRoom, code As String

    Set ex = ExampleRoomToRemove(Rooms)
    If ex Is Nothing Then Exit Sub
    code = LocValue(ex, "LOC_Code")

    If Not Back Is Nothing Then
        If StrComp(Back.Name, ex.Name, vbTextCompare) = 0 Then Set Back = Rooms.Item(1).Sheet
    End If

    wasQuiet = gQuiet
    gQuiet = True
    DeletePrintRoom ex
    gQuiet = wasQuiet

    ' A restored room called like the sample was created under "<name> 2".
    Dim rv As Variant
    For Each rv In Rooms
        Set room = rv
        If Not room.Sheet Is Nothing Then TidySheetName room.Sheet
    Next rv
    gQuiet = True
    RefreshLocations                ' the registry holds sheet names
    gQuiet = wasQuiet
    If SheetForCode(code) Is Nothing Then
        Summary = Summary & "- " & code & " (" & EXAMPLE_ROOM_NAME & "): sample print room removed" & vbCrLf
    Else
        Summary = Summary & "- " & code & " (" & EXAMPLE_ROOM_NAME & "): could not be removed - delete it with Remove print room" & vbCrLf
    End If
End Sub

' Finds the room by its code, creating it when the backup names a room this
' workbook does not have - the same duplicate-the-template path Add print
' room uses, so a restored room is built exactly like a hand-made one. Sets
' Room.Sheet (Nothing when the room is skipped) and Room.Created.
'
' Refresh Locations, at the end of CreatePrintRoom, announces itself with its
' own dialog; one per restored room would bury the restore's own summary, so
' that announcement is muted here (and still collected by QuietLog when a
' test is already running in quiet mode). The code is kept in full
' (KeepFullCode): a code Refresh Locations itself assigned can be longer than
' the 6 characters a typed one is capped at, and truncating it would stop the
' next restore of the same backup from finding the room it created.
Private Sub PrepareRoom(ByVal Room As clsRestoreRoom, ByRef Summary As String, ByRef Made As Long)
    Dim wasQuiet As Boolean, nm As String, sample As Worksheet, ok As Boolean

    If Len(Room.Code) = 0 Then
        Summary = Summary & "- (unknown location): skipped - the file does not say which print room it belongs to." & vbCrLf
        Exit Sub
    End If

    Set Room.Sheet = SheetForCode(Room.Code)
    If Not Room.Sheet Is Nothing Then
        Set sample = SampleRoomToReuse(Room)
        If sample Is Nothing Then Exit Sub

        ' The backup's room carries the sample's code: take over the sample's sheet.
        wasQuiet = gQuiet
        gQuiet = True
        ok = RecyclePrintRoom(sample, Room.RoomName, Room.Dept, Room.Code)
        gQuiet = wasQuiet
        If Not ok Then
            Set Room.Sheet = Nothing
            Summary = Summary & "- " & Room.Code & ": skipped - the sample print room could not be reused." & vbCrLf
            Exit Sub
        End If
        Room.Created = True
        Made = Made + 1
        ApplyRoomSettings Room.Sheet, Room.Settings
        Exit Sub
    End If

    nm = Room.RoomName
    If Len(nm) = 0 Then nm = Room.Code

    wasQuiet = gQuiet
    gQuiet = True
    Set Room.Sheet = CreatePrintRoom(nm, Room.Dept, Room.Code, True)
    gQuiet = wasQuiet

    If Room.Sheet Is Nothing Then
        Summary = Summary & "- " & Room.Code & ": skipped - the print room could not be created." & vbCrLf
        Exit Sub
    End If
    Room.Created = True
    Made = Made + 1
    ApplyRoomSettings Room.Sheet, Room.Settings
End Sub

' A new room starts from the template's defaults; when the source is an older
' workbook it carries the room's own, so the room comes back as it was set up.
' Written only to a room this restore has just created - an existing room's
' settings are left as they are, the way an existing room's job records are
' only added to - and BEFORE the job records go in, so the roll length unit is
' already the room's own when Import weighs a file's Qty unit against it (no
' conversion if they match).
'
' Events are off for the writes (AppOff), so changing LOC_RollUnit does not run
' the rescale of existing rows - there are none yet - and the hidden "unit the
' stored Qty values are in" marker is set by hand, exactly as the change
' handler would; the Qty validation message is reworded for the unit.
Private Sub ApplyRoomSettings(ByVal Ws As Worksheet, ByVal Settings As clsDict)
    Dim names As Variant, i As Long, c As Range, s As String, unit As String

    If Settings Is Nothing Then Exit Sub

    AppOff
    UnlockSheet Ws
    names = Array("LOC_DefTech", "LOC_DefPrinter", "LOC_DefPaper", "LOC_DefDisPaper", "LOC_DefDisCons", "LOC_Printers")
    For i = LBound(names) To UBound(names)
        If Settings.Exists(CStr(names(i))) Then
            s = Trim$(CStr(Settings.Item(CStr(names(i)))))
            Set c = LocRange(Ws, CStr(names(i)))
            If Not c Is Nothing Then
                If Len(s) > 0 Then c.Value = s
            End If
        End If
    Next i

    If Settings.Exists("LOC_RollUnit") Then
        unit = Trim$(CStr(Settings.Item("LOC_RollUnit")))
        If StrComp(unit, "Centimetres", vbTextCompare) = 0 Or StrComp(unit, "Metres", vbTextCompare) = 0 Then
            Set c = LocRange(Ws, "LOC_RollUnit")
            If Not c Is Nothing Then
                c.Value = unit
                SetAppliedRollUnit Ws, unit
            End If
        End If
    End If
    RelockSheet Ws
    RefreshQtyValidation Ws
    BindColumns Ws
    ClearBindAutofill Ws
    AppOn
End Sub

' A room with a single permitted printer has it filled into the table's one
' empty row by BindColumns (and that row stamped with its prices), which
' stops Import recognising the row as empty - the first job would be added
' beneath it and a phantom, dateless row left on top. The room is brand new,
' so its only row holds nothing worth keeping: blank it again.
Private Sub ClearBindAutofill(ByVal Ws As Worksheet)
    Dim lo As ListObject

    Set lo = JobsTable(Ws)
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count <> 1 Then Exit Sub
    If IsBlankRow(lo, 1) Then Exit Sub
    UnlockSheet Ws
    ClearTypedCells lo.ListRows(1).Range
    RelockSheet Ws
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
                             ByRef Appended As Long, ByRef Overwrite As Long, ByRef Skipped As Long, _
                             ByRef Renamed As Long)
    Dim keyHdr As String, d As clsDict, key As String, n As Long, v As Variant
    Dim notesVal As String, skip As Boolean
    Dim hasId As Boolean, idHdr As String, idCode As String, nameHdr As String, hwmKey As String

    keyHdr = CatalogKeyHeader(TableName)
    If Len(keyHdr) = 0 Then Exit Sub
    hasId = CatalogIdSpec(TableName, idHdr, idCode, nameHdr, hwmKey)

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
                If hasId Then RenameIfNameTaken lo, n, d, nameHdr, idCode, key, Renamed
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

    ' Rows written in bulk may carry numbers above anything this workbook has
    ' issued itself (a same-site backup restored into a newer workbook), so
    ' the ID counter is raised to cover them - see modCatalog "catalogue IDs".
    If hasId Then SyncCatalogHwm lo, TableName
End Sub

' Same-name-different-ID guard, for a restore that combines two workbooks'
' configurations. Every lookup in this workbook is by NAME (clsDict.Add
' replaces on a duplicate), so an incoming row whose name is already used by a
' DIFFERENT row would silently shadow it in every dropdown and cost lookup.
' Such a row is kept, under its own ID, with its site added to the name:
' "HP T730" becomes "HP T730 (SITE2)". The site is read off the incoming ID
' (SITE2-PRN-00007); an ID in some other shape falls back to "imported".
'
' SkipRow is the row the incoming one is about to overwrite (0 for an append):
' a row is never a clash with itself, which is also what makes restoring the
' same backup twice give the same result both times.
Private Sub RenameIfNameTaken(ByVal lo As ListObject, ByVal SkipRow As Long, ByVal d As clsDict, _
                              ByVal NameHdr As String, ByVal Code As String, ByVal Key As String, _
                              ByRef Renamed As Long)
    Dim nm As String, candidate As String, site As String, k As Long, p As Long

    If Not d.Exists(NameHdr) Then Exit Sub
    nm = Trim$(CStr(d.Item(NameHdr)))
    If Len(nm) = 0 Then Exit Sub
    If Not NameInUse(lo, NameHdr, nm, SkipRow) Then Exit Sub

    p = InStrRev(Key, "-" & Code & "-", -1, vbTextCompare)
    If p > 1 Then site = Left$(Key, p - 1) Else site = "imported"

    candidate = nm & " (" & site & ")"
    k = 1
    Do While NameInUse(lo, NameHdr, candidate, SkipRow)
        k = k + 1
        candidate = nm & " (" & site & " " & k & ")"
    Loop

    d.Add NameHdr, candidate
    Renamed = Renamed + 1
End Sub

Private Function NameInUse(ByVal lo As ListObject, ByVal NameHdr As String, ByVal Nm As String, _
                           ByVal SkipRow As Long) As Boolean
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If i <> SkipRow Then
            If StrComp(Trim$(CStr(CellIn(lo, i, NameHdr).Value)), Nm, vbTextCompare) = 0 Then
                NameInUse = True
                Exit Function
            End If
        End If
    Next i
End Function

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
        ' A column the backup does not have at all (an older backup restored
        ' over a newer workbook, e.g. a 1.3 backup's printers have no Colour
        ' mode or Template cost) is left exactly as it is, not cleared: that
        ' would silently turn a multi-pass printer back into a single-pass one
        ' (multi-pass design decision 34). A column the backup has but holds
        ' blank is still cleared.
        If Not cell.HasFormula And d.Exists(header) Then
            s = Trim$(CStr(d.Item(header)))
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
    PickBackupFile = PickCsvFile("Restore workbook - choose any one file from the backup, or an older copy of the workbook", _
        "Backup files (*.csv;*.xlsm;*.xlsx;*.xlsb),*.csv;*.xlsm;*.xlsx;*.xlsb")
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
