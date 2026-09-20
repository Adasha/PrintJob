Attribute VB_Name = "modExport"
Option Explicit

' Per-location export to CSV, and the knowledge of what has not been exported.
'
' Three jobs at once, which is why it earns its place in v1 rather than being
' a recovery mechanism nobody maintains:
'
'   1. a recovery path for a deleted print room sheet - Excel raises no event
'      when a sheet is deleted, so the records have to already be somewhere
'      else (design 3.3);
'   2. a way to get records out of the workbook for finance;
'   3. the transport a future aggregator reads (design 11.2).
'
' CONTENT IS THE IRREVERSIBLE DECISION. The file carries every column,
' including the snapshot block. An export without S_UnitCost, S_ConsRate and
' the rest can never be faithfully re-imported: re-importing it would recost
' every row at today's prices, destroying exactly the historical integrity the
' snapshot exists to protect. No version number rescues a value that was never
' written down.
'
' Column ORDER, by contrast, is free. The file carries a header row and any
' importer maps by name, so the job table can be reordered without
' invalidating older exports.

Private Const CSV_UTF8 As Long = 62         ' xlCSVUTF8
Private Const EXPORT_CELL As String = "$B$8"

' The canonical export order. Fixed here rather than mirroring the sheet, so
' files stay comparable across versions even after the table is reordered.
'
' H_Issues is deliberately absent: it is an internal working cell behind the
' Status message, meaningless outside the workbook and recomputed on import.
' Everything else is present, calculated columns included - they cost nothing
' and make the file readable by a person rather than only by a machine.
Private Function ExportColumns() As Variant
    ExportColumns = Array( _
        "Job ID", "Date/Time", "Student Name", "Student No", "Technician", _
        "Printer", "Paper Stock", "Unit", "Quantity", "Print Width mm", _
        "Disregard Paper", "Disregard Consumable", "Area m2", _
        "Paper Cost", "Consumable Cost", "Gross Cost", "Disregarded", _
        "Chargeable Cost", "Notes", "Status", _
        "S_PrinterID", "S_StockID", "S_TechID", "S_Family", "S_Measure", _
        "S_UnitCost", "S_StockWidth_mm", "S_SheetHeight_mm", "S_ConsRate", _
        "S_StampedAt", "S_StampedBy", "S_SchemaVer")
End Function

' =============================================================== export ===
Public Sub ExportLocation(ByVal ws As Worksheet)
    Dim n As Long, path As String, status As String

    On Error GoTo Fail
    status = ExportOne(ws, n, path)
    Select Case status
        Case "EMPTY"
            Say "There is nothing to export.", "'" & LocValue(ws, "LOC_Name") & "' has no print jobs recorded."
        Case "NOPATH"
            Say "This print room could not be exported.", _
                "The workbook's own folder could not be resolved to a location on this computer. " & _
                "That happens when the file is open from OneDrive and Excel reports its address as a web link rather than a folder.", _
                "Open the workbook from the OneDrive folder on this computer rather than from the browser, then try again."
        Case "NOFILE"
            Say "The export did not produce a file.", _
                "Excel accepted the save but nothing was written to:" & vbCrLf & path, _
                "Nothing has been marked as exported. Check you can write to that folder, then try again."
        Case ""
            Say n & " print job" & IIf(n = 1, "", "s") & " exported.", _
                "Written to:" & vbCrLf & path, _
                "The file holds every column, including the frozen prices each job was costed at, so it is a complete record of this print room."
    End Select
    Exit Sub
Fail:
    AppReset
    ReportError "Export"
End Sub

' Every registered print room, one CSV each, in one pass. Shares ExportOne
' with ExportLocation rather than looping the button macro, so this is one
' AppOff/AppOn bracket for the whole batch (no recalculation thrash between
' locations) and one summary dialog instead of N.
Public Sub ExportAllLocations()
    Dim ws As Worksheet, n As Long, path As String, status As String
    Dim done As Long, skipped As Long, detail As String, failed As String, why As String

    On Error GoTo Fail
    AppOff
    For Each ws In LocationSheets()
        status = ExportOne(ws, n, path)
        Select Case status
            Case ""
                done = done + 1
                detail = detail & "- " & LocValue(ws, "LOC_Name") & ": " & n & " job" & IIf(n = 1, "", "s") & vbCrLf
            Case "EMPTY"
                skipped = skipped + 1
            Case Else
                failed = failed & "- " & LocValue(ws, "LOC_Name") & ": " & status & vbCrLf
        End Select
    Next ws
    AppOn

    If done = 0 And skipped = 0 And Len(failed) = 0 Then
        Say "No print rooms were found.", "Run Refresh Locations first."
        Exit Sub
    End If

    why = detail
    If skipped > 0 Then why = why & skipped & " print room" & IIf(skipped = 1, "", "s") & " had nothing to export and " & IIf(skipped = 1, "was", "were") & " skipped." & vbCrLf
    If Len(failed) > 0 Then why = why & vbCrLf & "Could not be exported:" & vbCrLf & failed

    Say done & " print room" & IIf(done = 1, "", "s") & " exported.", why, _
        "Each file holds every column, including the frozen prices each job was costed at, so it is a complete record of that print room."
    Exit Sub
Fail:
    AppReset
    ReportError "Export All Locations"
End Sub

' Does the actual export, without showing anything. Returns "" on success
' (N and Path filled in, the location already stamped as exported), or one of
' EMPTY / NOPATH / NOFILE naming what stopped it - so ExportLocation and
' ExportAllLocations can each decide how to tell the user, one dialog at a
' time or rolled into a single summary.
Private Function ExportOne(ByVal ws As Worksheet, ByRef n As Long, ByRef path As String) As String
    Dim lo As ListObject, cols As Variant, block As Variant, wbOut As Workbook

    Set lo = JobsTable(ws)
    If lo Is Nothing Then
        ExportOne = "EMPTY"
        Exit Function
    End If
    n = RowCount(lo)
    If n = 0 Then
        ExportOne = "EMPTY"
        Exit Function
    End If

    cols = ExportColumns
    block = BuildBlock(ws, lo, n, cols)

    path = ExportPath(ws)
    If Len(path) = 0 Then
        ExportOne = "NOPATH"
        Exit Function
    End If

    AppOff
    Set wbOut = Application.Workbooks.Add
    With wbOut.Worksheets(1)
        ' Everything goes in as TEXT. Without this Excel re-parses what we
        ' carefully formatted: an ISO "2026-09-15 10:12:00" becomes a date
        ' value again and is written out in whatever the machine's locale
        ' prefers - "9/15/2026" on this one - and the schema version "1.0"
        ' becomes the number 1, so an importer checking for "1.0" sees "1".
        ' The staging sheet exists only to reach Excel's CSV writer; it must
        ' not interpret anything on the way.
        .Cells.NumberFormat = "@"
        .Range(.Cells(1, 1), .Cells(UBound(block, 1), UBound(block, 2))).Value = block
    End With
    ' Excel's own CSV writer, so quoting and comma escaping are its problem,
    ' not ours - and xlCSVUTF8 gets the pound sign and any non-ASCII student
    ' name right. VBA's Print # writes in the system codepage and mangles
    ' both; ADODB.Stream would, but does not exist on Mac.
    wbOut.SaveAs path, CSV_UTF8
    wbOut.Close False
    AppOn

    ' Confirm the file is actually there before claiming anything. SaveAs to a
    ' destination Excel accepts but does not write locally - a OneDrive URL, a
    ' path the CSV prompt was cancelled on - returns without raising, and the
    ' two lines below would then stamp the location as exported and tell the
    ' caller its records were safe. This status is what stands between a sheet
    ' deletion and the records, so it must never be optimistic.
    If Not FileExists(path) Then
        ExportOne = "NOFILE"
        Exit Function
    End If

    StampExported ws
    RefreshExportStatus ws
End Function

' Header block, then a blank line, then the header row, then the records.
Private Function BuildBlock(ByVal ws As Worksheet, ByVal lo As ListObject, _
                            ByVal n As Long, ByVal cols As Variant) As Variant
    Const HEAD As Long = 8
    Dim a() As Variant, w As Long, r As Long, c As Long, hdr As String
    w = UBound(cols) - LBound(cols) + 1
    If w < 2 Then w = 2
    ReDim a(1 To HEAD + 1 + n, 1 To w)

    a(1, 1) = "Print Cost Management export"
    a(2, 1) = "Schema version":  a(2, 2) = SCHEMA_VER
    a(3, 1) = "Site ID":         a(3, 2) = SettingText("SITE_ID", "SITE")
    a(4, 1) = "Site name":       a(4, 2) = SettingText("SITE_NAME")
    a(5, 1) = "Location code":   a(5, 2) = LocValue(ws, "LOC_Code")
    a(6, 1) = "Location name":   a(6, 2) = LocValue(ws, "LOC_Name")
    a(7, 1) = "Generated":       a(7, 2) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    a(8, 1) = "Rows":            a(8, 2) = n

    For c = LBound(cols) To UBound(cols)
        a(HEAD + 1, c - LBound(cols) + 1) = CStr(cols(c))
    Next c

    For r = 1 To n
        For c = LBound(cols) To UBound(cols)
            hdr = CStr(cols(c))
            a(HEAD + 1 + r, c - LBound(cols) + 1) = CellOut(lo, r, hdr)
        Next c
    Next r

    BuildBlock = a
End Function

' Dates are written as yyyy-mm-dd hh:nn:ss. A serial number would be unreadable
' and a locale-formatted date would be ambiguous on import - 06/07 is two
' different days depending on who opens it.
Private Function CellOut(ByVal lo As ListObject, ByVal RowNo As Long, ByVal Header As String) As Variant
    Dim c As Range, d As Double
    On Error GoTo Missing
    Set c = CellIn(lo, RowNo, Header)

    If StrComp(Header, "Date/Time", vbTextCompare) = 0 Or _
       StrComp(Header, "S_StampedAt", vbTextCompare) = 0 Then
        d = DateSerialOf(c)
        If d > 0 Then CellOut = Format$(CDate(d), "yyyy-mm-dd hh:nn:ss")
        Exit Function
    End If

    CellOut = c.Value2
    Exit Function
Missing:
    CellOut = ""
End Function

Private Function ExportPath(ByVal ws As Worksheet) As String
    Dim base As String, folder As String
    folder = ExportFolder()
    If Len(folder) = 0 Then Exit Function

    base = "PrintCosts-" & SettingText("SITE_ID", "SITE") & "-" & _
           LocValue(ws, "LOC_Code") & "-" & Format$(Now, "yyyymmdd-hhnn") & ".csv"
    ' Beside the workbook. No file dialog: Application.FileDialog does not
    ' exist on Mac, and a path the user already has open is one the Mac
    ' sandbox already permits.
    ExportPath = folder & Application.PathSeparator & base
End Function

' The folder to write into, as a real filesystem path.
'
' ThisWorkbook.Path is NOT reliably one. When a workbook is opened from a
' OneDrive-backed folder Excel may report its location as the service URL -
' "https://d.docs.live.net/<cid>/Development/.../src" - and this project lives
' in OneDrive. Concatenating that with Application.PathSeparator produced a
' hybrid like ".../src\PrintCosts-UNI-MAIN-....csv", SaveAs did not raise, and
' the export reported success having written nothing anywhere the user could
' find. Worse, it then stamped the location as exported, so the one warning
' standing between a sheet deletion and the records said the opposite of the
' truth.
'
' A local path is returned unchanged, which is the Mac case and the ordinary
' Windows case. A URL is mapped back onto the local OneDrive root.
Private Function ExportFolder() As String
    Dim p As String
    p = ThisWorkbook.path
    If Not IsUrl(p) Then
        ExportFolder = p
    Else
        ExportFolder = LocalRootOf(p)
    End If
End Function

Private Function IsUrl(ByVal p As String) As Boolean
    IsUrl = (InStr(1, p, "http://", vbTextCompare) = 1) Or _
            (InStr(1, p, "https://", vbTextCompare) = 1)
End Function

' Maps a OneDrive URL back to the synced folder on this machine.
'
' There is no API for this. The URL's leading segments vary by account type -
' a consumer CID, or a SharePoint site and library - so rather than parsing
' those, this drops the scheme and host and then tries progressively shorter
' tails of the remaining path against each local OneDrive root, taking the
' first that names a folder that actually exists. Returns "" when none does,
' which the caller reports rather than guessing.
Private Function LocalRootOf(ByVal url As String) As String
    Dim roots As Variant, parts As Variant
    Dim tail As String, cand As String, sep As String
    Dim i As Long, j As Long, k As Long

    roots = Array(Environ$("OneDrive"), Environ$("OneDriveConsumer"), Environ$("OneDriveCommercial"))
    sep = Application.PathSeparator

    tail = url
    i = InStr(tail, "://")
    If i > 0 Then tail = Mid$(tail, i + 3)
    i = InStr(tail, "/")
    If i = 0 Then Exit Function
    tail = Mid$(tail, i + 1)
    tail = Replace$(tail, "/", sep)

    parts = Split(tail, sep)
    For i = LBound(parts) To UBound(parts)
        cand = ""
        For j = i To UBound(parts)
            If Len(cand) > 0 Then cand = cand & sep
            cand = cand & CStr(parts(j))
        Next j
        For k = LBound(roots) To UBound(roots)
            If Len(CStr(roots(k))) > 0 Then
                If FolderExists(CStr(roots(k)) & sep & cand) Then
                    LocalRootOf = CStr(roots(k)) & sep & cand
                    Exit Function
                End If
            End If
        Next k
    Next i
End Function

Private Function FolderExists(ByVal p As String) As Boolean
    Dim a As Long
    On Error GoTo No
    a = GetAttr(p)
    FolderExists = ((a And vbDirectory) = vbDirectory)
    Exit Function
No:
    FolderExists = False
End Function

Private Function FileExists(ByVal p As String) As Boolean
    Dim a As Long
    On Error GoTo No
    a = GetAttr(p)
    FileExists = ((a And vbDirectory) <> vbDirectory)
    Exit Function
No:
    FileExists = False
End Function

' ========================================================== fingerprint ===
' What has changed since the last export is DERIVED, never tracked.
'
' A dirty flag would need setting by every mutator, and our own code runs with
' EnableEvents False so Worksheet_Change never fires for it. Worse, a flag
' lies in two situations nobody would think to test: a user can type into a
' sheet with macros disabled entirely, and restoring an old copy of a sheet
' brings its stale flag along with it. Both leave the workbook insisting a
' location was exported when it was not. A fingerprint computed on demand
' notices, the next time any code runs.
'
' The limitation, stated rather than hidden: a collision needs an edit that
' leaves the row count, the cell count, both sums and the latest date all
' unchanged. Two compensating changes in one location between exports.
' Constructible deliberately; vanishingly unlikely by accident.
Public Function ExportSig(ByVal ws As Worksheet) As String
    Dim lo As ListObject, n As Long
    Dim cells As Double, qty As Double, chg As Double, last As Double
    Dim bad As String

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Function
    n = RowCount(lo)
    If n = 0 Then
        ExportSig = "empty"
        Exit Function
    End If

    ' Each component is computed on its own. One blanket On Error Resume Next
    ' across all four is the pattern design 8.2 warns about: if ColIdx cannot
    ' find a column, execution skips to the next statement with that variable
    ' still 0, the remaining aggregates may be skipped too, and the fingerprint
    ' quietly loses a term. An edit that moves only that term would then read
    ' as "no changes since" - the one answer this function must never give
    ' wrongly, because it is what stands between a sheet deletion and the
    ' records. A component that cannot be computed is marked instead.
    cells = Agg(lo, "", "CountA", bad)
    qty = Agg(lo, "Quantity", "Sum", bad)
    chg = Agg(lo, "Chargeable Cost", "Sum", bad)
    last = Agg(lo, "Date/Time", "Max", bad)

    ExportSig = n & "|" & Format$(cells, "0") & "|" & Format$(qty, "0.####") & _
                "|" & Format$(chg, "0.####") & "|" & Format$(last, "0.######") & bad
End Function

' One aggregate over one column, or over the whole data range when Header is
' blank. On failure it appends a marker to Bad rather than returning a bare 0,
' so the caller's signature carries the fact that a term is missing instead of
' absorbing it. The marker is stable for a given fault, so a genuinely
' unchanged location still compares equal while the defect stays visible.
Private Function Agg(ByVal lo As ListObject, ByVal Header As String, _
                     ByVal What As String, ByRef Bad As String) As Double
    Dim rng As Range

    On Error GoTo Failed
    If Len(Header) = 0 Then
        Set rng = lo.DataBodyRange
    Else
        Set rng = lo.ListColumns(ColIdx(lo, Header)).DataBodyRange
    End If
    If rng Is Nothing Then GoTo Failed

    Select Case What
        Case "CountA": Agg = Application.WorksheetFunction.CountA(rng)
        Case "Sum":    Agg = Application.WorksheetFunction.Sum(rng)
        Case "Max":    Agg = Application.WorksheetFunction.Max(rng)
        Case Else:     GoTo Failed
    End Select
    Exit Function

Failed:
    Agg = 0
    Bad = Bad & "|?" & What
    If Len(Header) > 0 Then Bad = Bad & ":" & Header
End Function

' ============================================================= registry ===
Private Function RegRow(ByVal ws As Worksheet) As Long
    Dim lo As ListObject, i As Long
    Set lo = Tbl("tblLocations")
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "SheetName").Value)), ws.Name, vbTextCompare) = 0 Then
            RegRow = i
            Exit Function
        End If
    Next i
End Function

Public Sub StampExported(ByVal ws As Worksheet)
    Dim lo As ListObject, i As Long
    Set lo = Tbl("tblLocations")
    If lo Is Nothing Then Exit Sub
    i = RegRow(ws)
    If i = 0 Then Exit Sub
    UnlockSheet lo.Parent
    CellIn(lo, i, "Last export").Value = Now
    CellIn(lo, i, "Last export").NumberFormat = "dd/mm/yyyy hh:mm"
    CellIn(lo, i, "Export sig").Value = ExportSig(ws)
    RelockSheet lo.Parent
End Sub

' ================================================================ status ===
Public Function ExportStatusText(ByVal ws As Worksheet) As String
    Dim lo As ListObject, i As Long, was As String, now_ As String, whenDone As Variant
    Dim n As Long

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Function
    n = RowCount(lo)

    i = RegRow(ws)
    If i = 0 Then
        ExportStatusText = "Not yet registered - run Refresh Locations."
        Exit Function
    End If

    Set lo = Tbl("tblLocations")
    was = Trim$(CStr(CellIn(lo, i, "Export sig").Value))
    whenDone = CellIn(lo, i, "Last export").Value
    now_ = ExportSig(ws)

    If n = 0 Then
        ExportStatusText = "No print jobs recorded."
    ElseIf Len(was) = 0 Then
        ExportStatusText = n & " print job" & IIf(n = 1, "", "s") & " - never exported."
    ElseIf StrComp(was, now_, vbBinaryCompare) = 0 Then
        ExportStatusText = "Exported " & Format$(whenDone, "dd/mm/yyyy hh:mm") & " - no changes since."
    Else
        ExportStatusText = "CHANGED since the last export (" & Format$(whenDone, "dd/mm/yyyy hh:mm") & _
                           ") - export before deleting this sheet."
    End If
End Function

Public Function NeedsExport(ByVal ws As Worksheet) As Boolean
    Dim s As String
    s = ExportStatusText(ws)
    NeedsExport = (InStr(1, s, "CHANGED", vbBinaryCompare) > 0) Or _
                  (InStr(1, s, "never exported", vbTextCompare) > 0)
End Function

' Writes the status onto the sheet itself. Excel cannot intercept a sheet
' deletion, so the warning has to be somewhere the user is already looking
' when they right-click the tab.
Public Sub RefreshExportStatus(ByVal ws As Worksheet)
    Dim c As Range
    EnsureExportName ws
    Set c = LocRange(ws, "LOC_Export")
    If c Is Nothing Then Exit Sub

    UnlockSheet ws
    c.Offset(0, -1).Value = "Export"
    c.Offset(0, -1).Font.Bold = True
    c.Value = ExportStatusText(ws)
    c.Font.Bold = NeedsExport(ws)
    If NeedsExport(ws) Then
        c.Font.Color = RGB(176, 0, 32)
    Else
        c.Font.Color = RGB(90, 90, 90)
    End If
    c.Locked = True
    RelockSheet ws
End Sub

' The name is (re)created on every refresh rather than shipped in the .xlsx,
' so a duplicated or renamed sheet gets a correct one without hand surgery.
Private Sub EnsureExportName(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Names("LOC_Export").Delete
    On Error GoTo 0
    ws.Names.Add Name:="LOC_Export", RefersTo:="='" & ws.Name & "'!" & EXPORT_CELL
End Sub
