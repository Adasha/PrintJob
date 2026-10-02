Attribute VB_Name = "modExport"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

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
Private Const XLSX_FORMAT As Long = 51      ' xlOpenXMLWorkbook
' $AM$6, not $B$8 (2026-09-25 layout swap): Export moved to the side panel
' alongside Sheet status, trading places with the permitted-printers list
' and its "N printers (list)" display, which moved to the left - user
' request, both being read-only/derived state rather than day-to-day input.
' RefreshExportStatus's own Offset(0, -1) write for the "Export" label
' follows this automatically, onto AL6.
Private Const EXPORT_CELL As String = "$AM$6"

' Student names in the job-level CSV (design 10.4, O7). Export asks; a quiet
' (unattended) run cannot, so it includes names, as the backup and the tests
' always have, unless a test says otherwise through SetExportNames.
Private mQuietNames As Long

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
        "Printer", "Paper Stock", "Unit", "Qty", "Print Width mm", "Sheet size", _
        "Disregard Paper", "Disregard Consumable", "Area m2", _
        "Paper Cost", "Consumable Cost", "Gross Cost", "Disregarded", _
        "Chargeable Cost", "Paid", "Notes", "Status", _
        "S_PrinterID", "S_StockID", "S_TechID", "S_Measure", _
        "S_UnitCost", "S_StockWidth_mm", "S_SheetHeight_mm", "S_ConsRate", _
        "S_StampedAt", "S_StampedBy", "S_SchemaVer")
End Function

' =============================================================== export ===
Public Sub ExportLocation(ByVal ws As Worksheet)
    Dim n As Long, path As String, status As String, names As VbMsgBoxResult

    On Error GoTo Fail
    names = AskIncludeNames()
    If names = vbCancel Then Exit Sub
    status = ExportOne(ws, n, path, (names = vbYes))
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
                IIf(names = vbYes, "The file holds every column, including the frozen prices each job was costed at, so it is a complete record of this print room.", _
                    NamesOmittedNote())
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
' AskNames is False only for Backup, which is a recovery record and always
' keeps names.
Public Sub ExportAllLocations(Optional ByVal AskNames As Boolean = True)
    Dim ws As Worksheet, n As Long, path As String, status As String
    Dim done As Long, skipped As Long, detail As String, failed As String, why As String
    Dim names As VbMsgBoxResult

    On Error GoTo Fail
    names = vbYes
    If AskNames Then names = AskIncludeNames()
    If names = vbCancel Then Exit Sub
    AppOff
    For Each ws In LocationSheets()
        status = ExportOne(ws, n, path, (names = vbYes))
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
        IIf(names = vbYes, "Each file holds every column, including the frozen prices each job was costed at, so it is a complete record of that print room.", _
            NamesOmittedNote())
    Exit Sub
Fail:
    AppReset
    ReportError "Export All Locations"
End Sub

' Yes / No / Cancel: does this export carry student names? No blanks Student
' Name only - Student No stays, so a master can still total by student. The
' default button is No, the choice that sends less personal data out.
Private Function AskIncludeNames() As VbMsgBoxResult
    If gQuiet Then
        AskIncludeNames = IIf(mQuietNames = vbNo, vbNo, vbYes)
        Exit Function
    End If
    AskIncludeNames = MsgBox("Include student names in this export?" & vbCrLf & vbCrLf & _
        "Yes - the file is a complete record, and can restore a print room." & vbCrLf & _
        "No - student names are left blank; student numbers are kept." & vbCrLf & vbCrLf & _
        "Any names typed into the Notes column are not removed.", _
        vbQuestion + vbYesNoCancel + vbDefaultButton2, "Export")
End Function

Private Function NamesOmittedNote() As String
    NamesOmittedNote = "Student names were left out, so this is not a complete record of the print room. " & _
        "It has not been marked as exported, and a later import keeps the names already on the sheet."
End Function

' For tests only (called over COM): what a quiet run answers to the names
' question. Normal quiet behaviour, and the default, is Yes.
Public Sub SetExportNames(ByVal IncludeNames As Boolean)
    mQuietNames = IIf(IncludeNames, vbYes, vbNo)
End Sub

' Does the actual export, without showing anything. Returns "" on success
' (N and Path filled in, the location already stamped as exported), or one of
' EMPTY / NOPATH / NOFILE naming what stopped it - so ExportLocation and
' ExportAllLocations can each decide how to tell the user, one dialog at a
' time or rolled into a single summary.
Private Function ExportOne(ByVal ws As Worksheet, ByRef n As Long, ByRef path As String, _
                           Optional ByVal IncludeNames As Boolean = True) As String
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
    block = BuildBlock(ws, lo, n, cols, IncludeNames)

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

    ' A file without names cannot restore the room, so it must not clear the
    ' "unexported" warning that guards against a sheet deletion.
    If IncludeNames Then
        StampExported ws
        RefreshExportStatus ws
    End If
End Function

' ======================================================= report snapshot ===
' Problem 2's "Export report" - a formatted, point-in-time .xlsx archive of
' the Reports sheet's CURRENT filtered and sorted results, for human use and
' physical archiving.
'
' .xlsx over CSV or PDF. Requirements were: enough formatting to style for
' readability (CSV fails outright), opens directly and consistently on most
' computers, modifiable with ordinary software (PDF needs dedicated tools to
' change), and machine-readable text with no OCR needed - this project has
' already been burned once by a PDF whose text layer was not actually
' extractable (see docs/HANDOFF up to stage 7.md, the removed design-doc
' PDF). Written as static VALUES, not live formulas - an archive snapshot of
' what the filters showed at the moment of export, same as modExport's own
' CSVs are a snapshot rather than a link back into the workbook.
Public Sub ExportReportSnapshot(ByVal repWs As Worksheet)
    Const HDR_ROW As Long = 15
    Dim lastCol As Long, rng As Range, block As Variant, n As Long
    Dim path As String, wbOut As Workbook
    Dim promoted As Collection, header As Variant, full As Variant, tableHeaderRow As Long
    Dim totalChargeable As Double, stillOwed As Double
    Dim footer As Variant, footerStartRow As Long

    On Error GoTo Fail
    lastCol = LastHeaderColumn(repWs, HDR_ROW)
    If lastCol = 0 Then
        Say "There is nothing to export.", "No result columns were found on the Reports sheet."
        Exit Sub
    End If

    On Error Resume Next
    Set rng = repWs.Range("A16").SpillingToRange
    On Error GoTo 0
    If rng Is Nothing Then
        Say "There is nothing to export.", "The Reports sheet has no results yet."
        Exit Sub
    End If
    ' A16 may be spilling FILTER's own "no jobs match"/"no jobs recorded"
    ' fallback TEXT rather than real rows - the first cell of a real row is
    ' always a Date/Time serial number, so ISNUMBER is what tells them apart.
    If Not IsNumeric(rng.Cells(1, 1).Value2) Then
        Say "There is nothing to export.", CStr(rng.Cells(1, 1).Value)
        Exit Sub
    End If
    n = rng.Rows.Count

    block = SnapshotBlock(repWs, rng, HDR_ROW, lastCol, n)
    ' "Export names" (Reports!O10): the live results always show student
    ' name/no; only the exported file honours the toggle, so blank them here
    ' unless it is Yes. Before PromoteUniformColumns, which leaves an
    ' all-blank column alone.
    block = BlankNameColumns(block, repWs)
    block = RemoveExcludedColumns(block)
    ' Snag list items 2b/2c: a field that holds the SAME value on every
    ' exported row is a fact about the whole report, not a per-row detail -
    ' promoted into the header block and dropped from the table so the file
    ' reads as "this report is about X" rather than repeating X down a whole
    ' column. Live-sheet-only per the user's own answer on review: this
    ' happens here, at export time, never to the Reports sheet itself.
    block = PromoteUniformColumns(block, promoted)

    ' Total chargeable / Still owed reuse the Reports sheet's own "Matching"
    ' totals (F13/O13, modReports.BuildReports) rather than re-summing here -
    ' those are already computed over the exact same filter criteria the
    ' export is a snapshot of, so there is exactly one place that knows how
    ' "still owed" reconciles to "total chargeable minus paid".
    totalChargeable = SafeNum(repWs.Range("F13").Value)
    stillOwed = SafeNum(repWs.Range("O13").Value)

    block = AppendTotalsRow(block)
    header = SnapshotHeaderBlock(repWs, rng, n, promoted, totalChargeable, stillOwed)
    full = CombineBlocks(header, block)
    tableHeaderRow = UBound(header, 1) + 2   ' + 1 blank separator + 1 to reach the header row itself

    ' Schema version / Generated sit below the table, not above it - the
    ' machine-facing provenance a person reads once, if ever, rather than the
    ' report-about-what a person reads first. footerStartRow is the same
    ' "+1 blank separator, +1 to reach the row itself" arithmetic as
    ' tableHeaderRow above, measured from wherever the table actually ends.
    footer = SnapshotFooterBlock()
    footerStartRow = UBound(full, 1) + 2
    full = CombineBlocks(full, footer)

    path = ReportSnapshotPath()
    If Len(path) = 0 Then
        Say "This report could not be exported.", _
            "The workbook's own folder could not be resolved to a location on this computer. " & _
            "That happens when the file is open from OneDrive and Excel reports its address as a web link rather than a folder.", _
            "Open the workbook from the OneDrive folder on this computer rather than from the browser, then try again."
        Exit Sub
    End If

    AppOff
    Set wbOut = Application.Workbooks.Add
    With wbOut.Worksheets(1)
        .Range(.Cells(1, 1), .Cells(UBound(full, 1), UBound(full, 2))).Value = full
        .Range(.Cells(1, 1), .Cells(1, 2)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 2)).Font.Size = 14
        ' Total chargeable / Still owed, rows 8 and 9 of the header block -
        ' see SnapshotHeaderBlock. Bolded and currency-formatted the same way
        ' the on-sheet Matching totals (F13/O13) already are, so the numbers
        ' read as money rather than bare decimals.
        .Range(.Cells(8, 1), .Cells(9, 2)).Font.Bold = True
        .Range(.Cells(8, 2), .Cells(9, 2)).NumberFormat = CurrencyFormatCode()
        .Rows(tableHeaderRow).Font.Bold = True
        .Rows(tableHeaderRow).Interior.Color = RGB(222, 232, 244)
        FormatSnapshotColumns wbOut.Worksheets(1), block
        ' The totals row appended by AppendTotalsRow - last row of the table.
        .Rows(tableHeaderRow + UBound(block, 1) - 1).Font.Bold = True
        .Rows(tableHeaderRow + UBound(block, 1) - 1).Borders(xlEdgeTop).Weight = xlThin
        ' Schema version / Generated, below the table - de-emphasised the same
        ' grey already used for read-only/reference text elsewhere in the
        ' workbook (modReports.LegendRow's GreyText, CritCell's hint text),
        ' so provenance reads as background information, not as part of the
        ' report itself.
        .Range(.Cells(footerStartRow, 1), .Cells(footerStartRow + UBound(footer, 1) - 1, 2)).Font.Color = RGB(110, 110, 110)
        .Columns.AutoFit
    End With
    wbOut.SaveAs path, XLSX_FORMAT
    wbOut.Close False
    AppOn

    If Not FileExists(path) Then
        Say "The export did not produce a file.", _
            "Excel accepted the save but nothing was written to:" & vbCrLf & path, _
            "Nothing has been marked as exported. Check you can write to that folder, then try again."
        Exit Sub
    End If

    StampReportsExport repWs

    Say n & " record" & IIf(n = 1, "", "s") & " exported.", _
        "Written to:" & vbCrLf & path, _
        "A point-in-time copy of the filtered, sorted results currently shown, for human use or physical archiving. " & _
        "It will not update - run Export report again after changing the filters."
    Exit Sub
Fail:
    AppReset
    ReportError "Export report"
End Sub

' The last column with a header, excluding the trailing Job ID correlation
' column - an internal key, not part of the human-facing archive.
'
' Deliberately NOT filtered by Hidden (renamed from LastVisibleColumn,
' 2026-09-22, snag list items 2b/2c): the Reports sheet's minimum-columns
' view (2d) now hides several columns - Printer, Technician among them - that
' are still needed here as PromoteUniformColumns candidates. The old
' Hidden-filtered version silently truncated every export at the last
' visible column, dropping Technician and Notes from the file entirely
' without any error - found while building the promotion logic, not
' reported by a user, since nothing about a short file looks wrong until you
' already know a column is missing.
Private Function LastHeaderColumn(ByVal ws As Worksheet, ByVal HdrRow As Long) As Long
    Dim c As Long, last As Long, hdr As String
    c = 1
    Do While Len(Trim$(CStr(ws.Cells(HdrRow, c).Value))) > 0
        hdr = CStr(ws.Cells(HdrRow, c).Value)
        If StrComp(hdr, "Job ID", vbTextCompare) <> 0 Then last = c
        c = c + 1
        If c > 100 Then Exit Do
    Loop
    LastHeaderColumn = last
End Function

' Columns left out of the exported table by default - just Area m2 for now,
' since it is derivable from Qty and the paper's own dimensions and isn't
' something anyone reads directly off an export. Not user-configurable yet:
' a fixed list here, same status as ApplyReportsMinimumColumns' own "keep"
' set on the live Reports sheet (modReports) - a documented future
' enhancement is letting someone choose which columns an export includes,
' rather than a Settings toggle built now.
Private Function ExcludedExportColumns() As Variant
    ExcludedExportColumns = Array("Area m2")
End Function

' Blanks the Student name / Student no VALUES (not the columns) of BLOCK
' (header row 1, data rows 2..) unless the Reports sheet's Export names
' toggle (O10) is exactly "Yes" - so anything else, including an empty cell,
' is treated as No (data protection: opt in to reveal, not opt out).
' Matches by header text, like the rest of this module, so it does not care
' where the columns sit.
Private Function BlankNameColumns(ByVal block As Variant, ByVal repWs As Worksheet) As Variant
    Dim c As Long, r As Long, hdr As String

    If StrComp(Trim$(CStr(repWs.Range("O10").Value)), "Yes", vbTextCompare) = 0 Then
        BlankNameColumns = block
        Exit Function
    End If
    For c = 1 To UBound(block, 2)
        hdr = CStr(block(1, c))
        If StrComp(hdr, "Student name", vbTextCompare) = 0 Or StrComp(hdr, "Student no", vbTextCompare) = 0 Then
            For r = 2 To UBound(block, 1)
                block(r, c) = ""
            Next r
        End If
    Next c
    BlankNameColumns = block
End Function

Private Function IsExcludedColumn(ByVal Header As String, ByVal excluded As Variant) As Boolean
    Dim i As Long
    For i = LBound(excluded) To UBound(excluded)
        If StrComp(Header, CStr(excluded(i)), vbTextCompare) = 0 Then
            IsExcludedColumn = True
            Exit Function
        End If
    Next i
End Function

' Drops the columns ExcludedExportColumns names and closes the gap, rather
' than leaving a blank column behind - the same "keep" boolean-array
' technique PromoteUniformColumns (below) already uses to drop a column.
' Runs before PromoteUniformColumns: a column dropped here is simply gone,
' never a promotion candidate either.
Private Function RemoveExcludedColumns(ByVal block As Variant) As Variant
    Dim excluded As Variant, c As Long, r As Long
    Dim rows As Long, cols As Long, keep() As Boolean, nKeep As Long, outCol As Long
    Dim out() As Variant

    excluded = ExcludedExportColumns()
    rows = UBound(block, 1)
    cols = UBound(block, 2)
    ReDim keep(1 To cols)
    nKeep = 0
    For c = 1 To cols
        keep(c) = Not IsExcludedColumn(CStr(block(1, c)), excluded)
        If keep(c) Then nKeep = nKeep + 1
    Next c
    If nKeep = cols Then
        RemoveExcludedColumns = block
        Exit Function
    End If

    ReDim out(1 To rows, 1 To nKeep)
    For r = 1 To rows
        outCol = 0
        For c = 1 To cols
            If keep(c) Then
                outCol = outCol + 1
                out(r, outCol) = block(r, c)
            End If
        Next c
    Next r
    RemoveExcludedColumns = out
End Function

' The financial columns totalled by AppendTotalsRow below - the same set
' FormatSnapshotColumns already currency-formats, minus nothing: every money
' column in the results table gets a total, not a curated subset.
Private Function MoneyHeaders() As Variant
    MoneyHeaders = Array("Paper cost", "Consumable cost", "Gross", "Disregarded", "Chargeable")
End Function

Private Function IsMoneyHeader(ByVal Header As String, ByVal candidates As Variant) As Boolean
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If StrComp(Header, CStr(candidates(i)), vbTextCompare) = 0 Then
            IsMoneyHeader = True
            Exit Function
        End If
    Next i
End Function

' Appends one row to the bottom of BLOCK (header row 1, data rows 2..) summing
' each financial column - Qty/Area m2 are numeric too but are not money, so
' they are deliberately left blank on this row rather than summed. Runs AFTER
' PromoteUniformColumns, so it totals exactly the columns that end up in the
' file, whichever of the promotable ones (Location, Printer, ...) survived.
Private Function AppendTotalsRow(ByVal block As Variant) As Variant
    Dim rows As Long, cols As Long, r As Long, c As Long, hdr As String, s As Double
    Dim candidates As Variant, out() As Variant

    candidates = MoneyHeaders()
    rows = UBound(block, 1)
    cols = UBound(block, 2)

    ReDim out(1 To rows + 1, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols
            out(r, c) = block(r, c)
        Next c
    Next r

    out(rows + 1, 1) = "Total"
    For c = 1 To cols
        hdr = CStr(block(1, c))
        If IsMoneyHeader(hdr, candidates) Then
            s = 0
            For r = 2 To rows
                If IsNumeric(block(r, c)) Then s = s + CDbl(block(r, c))
            Next r
            out(rows + 1, c) = s
        End If
    Next c

    AppendTotalsRow = out
End Function

Private Function SafeNum(ByVal v As Variant) As Double
    If IsNumeric(v) Then SafeNum = CDbl(v)
End Function

' Header candidates for single-value promotion (2b/2c). Student name/no are
' included even though the Export names toggle usually blanks them - when the toggle is
' off every value is blank, so PromoteUniformColumns' own "at least one
' non-blank value" rule already leaves them alone with no special-casing
' needed here.
Private Function CandidatePromotionHeaders() As Variant
    CandidatePromotionHeaders = Array("Student name", "Student no", "Location", _
        "Printer", "Paper stock", "Technician")
End Function

Private Function IsPromotionCandidate(ByVal Header As String, ByVal candidates As Variant) As Boolean
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If StrComp(Header, CStr(candidates(i)), vbTextCompare) = 0 Then
            IsPromotionCandidate = True
            Exit Function
        End If
    Next i
End Function

' Scans BLOCK (header row 1, data rows 2..) for candidate columns whose
' non-blank values are all identical, and lifts each one out: added to
' Promoted as a "Header: Value" line for the header block, and dropped from
' the returned table so the gap closes rather than leaving an empty column.
' A candidate column with no non-blank values at all is left exactly where
' it is - there is nothing true to state about it, and removing an
' all-blank column would look like data loss rather than tidying.
Private Function PromoteUniformColumns(ByVal block As Variant, ByRef promoted As Collection) As Variant
    Dim candidates As Variant, c As Long, r As Long
    Dim hdr As String, v As String, uniform As String
    Dim rows As Long, cols As Long, keep() As Boolean, nKeep As Long, outCol As Long
    Dim hasValue As Boolean, mismatch As Boolean
    Dim out() As Variant

    Set promoted = New Collection
    candidates = CandidatePromotionHeaders()
    rows = UBound(block, 1)
    cols = UBound(block, 2)
    ReDim keep(1 To cols)
    For c = 1 To cols
        keep(c) = True
    Next c

    For c = 1 To cols
        hdr = CStr(block(1, c))
        If IsPromotionCandidate(hdr, candidates) Then
            hasValue = False
            mismatch = False
            uniform = ""
            For r = 2 To rows
                v = Trim$(CStr(block(r, c)))
                If Len(v) > 0 Then
                    If Not hasValue Then
                        uniform = v
                        hasValue = True
                    ElseIf StrComp(v, uniform, vbBinaryCompare) <> 0 Then
                        mismatch = True
                        Exit For
                    End If
                End If
            Next r
            If hasValue And Not mismatch Then
                promoted.Add hdr & ": " & uniform
                keep(c) = False
            End If
        End If
    Next c

    nKeep = 0
    For c = 1 To cols
        If keep(c) Then nKeep = nKeep + 1
    Next c
    If nKeep = cols Then
        PromoteUniformColumns = block
        Exit Function
    End If

    ReDim out(1 To rows, 1 To nKeep)
    For r = 1 To rows
        outCol = 0
        For c = 1 To cols
            If keep(c) Then
                outCol = outCol + 1
                out(r, outCol) = block(r, c)
            End If
        Next c
    Next r
    PromoteUniformColumns = out
End Function

' The metadata block written above the results table: what report this is,
' the site and date range it covers, how many rows, the Chargeable/Still owed
' summary, then one line per field PromoteUniformColumns lifted out. Schema
' version and Generated are provenance rather than report content, so they
' live in SnapshotFooterBlock, below the table, instead.
'
' Blank rows after the title and after Rows are deliberate breathing room -
' one groups "what this report is" (title), the next "what it covers" (site,
' date range, row count), the next "what it totals to" (chargeable/owed) -
' rather than nine lines running together as one undifferentiated list.
Private Function SnapshotHeaderBlock(ByVal repWs As Worksheet, ByVal rng As Range, _
                                     ByVal n As Long, ByVal promoted As Collection, _
                                     ByVal TotalChargeable As Double, ByVal StillOwed As Double) As Variant
    Const FIXED As Long = 9
    Dim a() As Variant, r As Long, item As Variant

    ReDim a(1 To FIXED + promoted.Count, 1 To 2)
    a(1, 1) = "Print job report"
    ' Row 2 left blank - gap below the title.
    a(3, 1) = "Site ID":        a(3, 2) = SettingText("SITE_ID", "SITE")
    a(4, 1) = "Site name":      a(4, 2) = SettingText("SITE_NAME")
    a(5, 1) = "Date range":     a(5, 2) = DateRangeText(repWs, rng)
    a(6, 1) = "Rows":           a(6, 2) = n
    ' Row 7 left blank - gap above the Chargeable/Still owed summary.
    ' Snag list item 1c's "total minus paid" reconciliation, restated here for
    ' the archive: Still owed is the same Chargeable total with paid charges
    ' removed, not a separate figure that could drift from it.
    a(8, 1) = "Total chargeable": a(8, 2) = TotalChargeable
    a(9, 1) = "Still owed":       a(9, 2) = StillOwed

    r = FIXED
    For Each item In promoted
        r = r + 1
        a(r, 1) = CStr(item)
    Next item

    SnapshotHeaderBlock = a
End Function

' Provenance, below the table rather than above it - schema version and the
' timestamp are for whatever machine or person re-reads this file later, not
' part of the report a person opens it to read first. Font.Color is set by
' the caller (ExportReportSnapshot), the same low-contrast grey already used
' for read-only/reference text elsewhere in the workbook.
Private Function SnapshotFooterBlock() As Variant
    Dim a(1 To 2, 1 To 2) As Variant
    a(1, 1) = "Schema version": a(1, 2) = SCHEMA_VER
    a(2, 1) = "Generated":      a(2, 2) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    SnapshotFooterBlock = a
End Function

' The From/To filter boxes (B7/B8) name the range the user actually asked
' for, so they win when set - even a one-sided filter ("From 01/01/2026,
' no end") is more informative than a min/max over whatever happened to
' match. Only once neither is set does this fall back to the actual spread
' of the exported rows' own Date/Time column.
Private Function DateRangeText(ByVal repWs As Worksheet, ByVal rng As Range) As String
    Dim fromD As Variant, toD As Variant, minD As Double, maxD As Double

    fromD = repWs.Range("B7").Value
    toD = repWs.Range("B8").Value
    If IsDate(fromD) Or IsDate(toD) Then
        If IsDate(fromD) And IsDate(toD) Then
            DateRangeText = Format$(CDate(fromD), "dd/mm/yyyy") & " to " & Format$(CDate(toD), "dd/mm/yyyy")
        ElseIf IsDate(fromD) Then
            DateRangeText = "From " & Format$(CDate(fromD), "dd/mm/yyyy")
        Else
            DateRangeText = "Up to " & Format$(CDate(toD), "dd/mm/yyyy")
        End If
        Exit Function
    End If

    minD = Application.WorksheetFunction.Min(rng.Columns(1))
    maxD = Application.WorksheetFunction.Max(rng.Columns(1))
    If Format$(CDate(minD), "yyyy-mm-dd") = Format$(CDate(maxD), "yyyy-mm-dd") Then
        DateRangeText = Format$(CDate(minD), "dd/mm/yyyy")
    Else
        DateRangeText = Format$(CDate(minD), "dd/mm/yyyy") & " to " & Format$(CDate(maxD), "dd/mm/yyyy")
    End If
End Function

' Stacks the header block above the results table with one blank row between
' them, widening to whichever of the two is wider (the header block is only
' ever two columns; the table is usually more) so neither side is clipped.
Private Function CombineBlocks(ByVal top As Variant, ByVal bottom As Variant) As Variant
    Dim topRows As Long, topCols As Long, botRows As Long, botCols As Long
    Dim w As Long, r As Long, c As Long
    Dim out() As Variant

    topRows = UBound(top, 1)
    topCols = UBound(top, 2)
    botRows = UBound(bottom, 1)
    botCols = UBound(bottom, 2)
    w = topCols
    If botCols > w Then w = botCols

    ReDim out(1 To topRows + 1 + botRows, 1 To w)
    For r = 1 To topRows
        For c = 1 To topCols
            out(r, c) = top(r, c)
        Next c
    Next r
    For r = 1 To botRows
        For c = 1 To botCols
            out(topRows + 1 + r, c) = bottom(r, c)
        Next c
    Next r
    CombineBlocks = out
End Function

Private Function SnapshotBlock(ByVal ws As Worksheet, ByVal rng As Range, _
                               ByVal HdrRow As Long, ByVal LastCol As Long, ByVal n As Long) As Variant
    Dim a() As Variant, r As Long, c As Long
    ReDim a(1 To n + 1, 1 To LastCol)
    For c = 1 To LastCol
        a(1, c) = ws.Cells(HdrRow, c).Value
    Next c
    For r = 1 To n
        For c = 1 To LastCol
            a(r + 1, c) = rng.Cells(r, c).Value
        Next c
    Next r
    SnapshotBlock = a
End Function

Private Sub FormatSnapshotColumns(ByVal outWs As Worksheet, ByVal block As Variant)
    Dim c As Long, hdr As String
    For c = 1 To UBound(block, 2)
        hdr = CStr(block(1, c))
        Select Case hdr
            Case "Date/Time"
                outWs.Columns(c).NumberFormat = "dd/mm/yyyy hh:mm"
            Case "Paper cost", "Consumable cost", "Gross", "Disregarded", "Chargeable"
                outWs.Columns(c).NumberFormat = CurrencyFormatCode()
            Case "Qty", "Area m2"
                outWs.Columns(c).NumberFormat = "#,##0.00"
        End Select
    Next c
End Sub

Private Function ReportSnapshotPath() As String
    Dim folder As String, base As String
    folder = ExportFolder()
    If Len(folder) = 0 Then Exit Function
    base = "PrintCosts-Report-" & SettingText("SITE_ID", "SITE") & "-" & Format$(Now, "yyyymmdd-hhnn") & ".xlsx"
    ReportSnapshotPath = folder & Application.PathSeparator & base
End Function

' Header block, then a blank line, then the header row, then the records.
Private Function BuildBlock(ByVal ws As Worksheet, ByVal lo As ListObject, _
                            ByVal n As Long, ByVal cols As Variant, _
                            Optional ByVal IncludeNames As Boolean = True) As Variant
    Const HEAD As Long = 8
    Dim a() As Variant, w As Long, r As Long, c As Long, hdr As String
    w = UBound(cols) - LBound(cols) + 1
    If w < 4 Then w = 4
    ReDim a(1 To HEAD + 1 + n, 1 To w)

    a(1, 1) = "Print Cost Management export"
    a(2, 1) = "Schema version":  a(2, 2) = SCHEMA_VER
    a(3, 1) = "Site ID":         a(3, 2) = SettingText("SITE_ID", "SITE")
    a(4, 1) = "Site name":       a(4, 2) = SettingText("SITE_NAME")
    a(5, 1) = "Location code":   a(5, 2) = LocValue(ws, "LOC_Code")
    a(6, 1) = "Location name":   a(6, 2) = LocValue(ws, "LOC_Name")
    a(7, 1) = "Generated":       a(7, 2) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    a(8, 1) = "Rows":            a(8, 2) = n
    ' Beside Rows, not a new line: the header row stays on row 9, where every
    ' importer already looks for it. Import reads this to leave names alone.
    If Not IncludeNames Then a(8, 3) = "Student names": a(8, 4) = "Omitted"

    For c = LBound(cols) To UBound(cols)
        a(HEAD + 1, c - LBound(cols) + 1) = CStr(cols(c))
    Next c

    For r = 1 To n
        For c = LBound(cols) To UBound(cols)
            hdr = CStr(cols(c))
            If Not IncludeNames And StrComp(hdr, "Student Name", vbTextCompare) = 0 Then
                a(HEAD + 1 + r, c - LBound(cols) + 1) = ""
            Else
                a(HEAD + 1 + r, c - LBound(cols) + 1) = CellOut(lo, r, hdr)
            End If
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
' EXPORT_FOLDER (snag list item 10) overrides the default when set and the
' path actually exists on this computer - checked rather than trusted, so a
' folder that only existed on whoever set it does not silently swallow every
' export on a different machine. It falls through to the workbook-relative
' default in that case rather than exporting nowhere.
'
' Otherwise: ThisWorkbook.Path is NOT reliably a filesystem path. When a
' workbook is opened from a OneDrive-backed folder Excel may report its
' location as the service URL -
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
' Public (2026-09-22, snag 4a): modBackup.BackupAll needs the same resolved
' folder for its catalogue-table CSVs as the per-location exports already
' use, and the OneDrive-URL resolution logic below is exactly the kind of
' thing that must never be duplicated (§9.3's own env-var gotcha lives here).
Public Function ExportFolder() As String
    Dim custom As String, p As String
    custom = Trim$(SettingText("EXPORT_FOLDER"))
    If Len(custom) > 0 Then
        If FolderExists(custom) Then
            ExportFolder = custom
            Exit Function
        End If
    End If

    p = ThisWorkbook.path
    If Not IsUrl(p) Then
        ExportFolder = p
    Else
        ExportFolder = LocalRootOf(p)
        If Len(ExportFolder) = 0 Then ExportFolder = FallbackFolder()
    End If
End Function

' Last resort when the OneDrive URL cannot be mapped to a folder this machine
' can see (e.g. a SharePoint library opened online rather than synced).
' Mac: the sandbox's own Documents folder, which Excel can always write.
' Windows: the user's Documents folder. The success message names the full
' path, so the file is never lost, just not beside the workbook.
Private Function FallbackFolder() As String
    Dim h As String, d As String
    h = Environ$("HOME")
    If Len(h) > 0 Then
        d = h & "/Documents"
    Else
        h = Environ$("USERPROFILE")
        If Len(h) = 0 Then Exit Function
        d = h & "\Documents"
    End If
    If FolderExists(d) Then FallbackFolder = d
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
    Dim roots As Collection, parts As Variant, root As Variant
    Dim tail As String, cand As String, sep As String
    Dim i As Long, j As Long

    Set roots = CandidateOneDriveRoots()
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
        For Each root In roots
            If FolderExists(CStr(root) & sep & cand) Then
                LocalRootOf = CStr(root) & sep & cand
                Exit Function
            End If
        Next root
    Next i
End Function

' Windows sets OneDrive/OneDriveConsumer/OneDriveCommercial environment
' variables naming the sync root(s) directly. Mac never has - the OneDrive
' Mac client sets no equivalent - and instead syncs each account under
' ~/Library/CloudStorage/OneDrive-<AccountName>, so that folder is scanned
' for candidates instead. This is what let a Mac user reproduce the NOPATH
' error even when the workbook's OneDrive folder was fully synced locally:
' the Windows-only env vars were the only roots ever tried.
Private Function CandidateOneDriveRoots() As Collection
    Dim out As New Collection, v As Variant
    Dim home As String, base As String, name As String

    For Each v In Array(Environ$("OneDrive"), Environ$("OneDriveConsumer"), Environ$("OneDriveCommercial"))
        If Len(CStr(v)) > 0 Then out.Add CStr(v)
    Next v

    home = RealMacHome()
    If Len(home) > 0 Then
        base = home & "/Library/CloudStorage"
        name = Dir$(base & "/OneDrive*", vbDirectory)
        Do While Len(name) > 0
            If name <> "." And name <> ".." Then out.Add base & "/" & name
            name = Dir$()
        Loop
        ' The older OneDrive client layout, still used by some installs.
        If FolderExists(home & "/OneDrive") Then out.Add home & "/OneDrive"
    End If

    Set CandidateOneDriveRoots = out
End Function

' Excel for Mac is sandboxed, and inside the sandbox HOME is the app's
' container (~/Library/Containers/com.microsoft.Excel/Data), not the user's
' home - so "$HOME/Library/CloudStorage" named a folder that does not exist
' and the scan above found nothing. The real home is whatever precedes
' "/Library/Containers/". Returns "" off Mac.
Private Function RealMacHome() As String
    Dim h As String, i As Long
    h = Environ$("HOME")
    If Len(h) = 0 Then Exit Function
    i = InStr(1, h, "/Library/Containers/", vbTextCompare)
    If i > 1 Then h = Left$(h, i - 1)
    RealMacHome = h
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
    qty = Agg(lo, "Qty", "Sum", bad)
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
