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

' The seven catalogue/configuration tables this backs up. Settings is one of
' them, not a separate mechanism - modUtils.Tbl finds it on the Settings
' sheet exactly like the other three lookup tables there.
Private Function CatalogTableNames() As Variant
    CatalogTableNames = Array("tblTechnicians", "tblPrinters", "tblPapers", _
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
    WriteBlockCsv = FileExists(path)
    Exit Function
Failed:
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close False
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
        "A print room listed above that does not exist here yet will be CREATED. " & _
        "Job records will be imported into each print room listed above, following the same rule as Import (existing Job ID: overwritten; new Job ID: added). " & _
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
    LoadRestorePlan = LoadFromCsvSet(AnyFilePath, CatalogRows, Rooms, SrcDesc)
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
    Dim wasQuiet As Boolean, nm As String

    If Len(Room.Code) = 0 Then
        Summary = Summary & "- (unknown location): skipped - the file does not say which print room it belongs to." & vbCrLf
        Exit Sub
    End If

    Set Room.Sheet = SheetForCode(Room.Code)
    If Not Room.Sheet Is Nothing Then Exit Sub

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
    PickBackupFile = PickCsvFile("Restore workbook - choose any one file from the backup")
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
