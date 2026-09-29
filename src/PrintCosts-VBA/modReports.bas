Attribute VB_Name = "modReports"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Builds the Summary and Reports sheets.
'
' This module writes LAYOUT AND FORMULAS ONCE. It does no reporting: every
' figure on both sheets is a live worksheet formula over the consolidated
' range on _Data, so the reports cannot go stale and there is no refresh
' button to forget. The design document's claim that "modReports does not
' exist" was half right - what does not exist is a VBA reporting engine.
'
' Both builders are idempotent. They are re-run by InitialiseWorkbook, which
' is also what picks up a change to the job table's columns.
'
' Columns are addressed by HEADER NAME, never by position:
'
'     INDEX(_Data!$A$10#,,MATCH("Qty",_Data!$A$9:$AZ$9,0))
'
' It is wordier than INDEX(...,,10) and it is the reason the job table's
' columns can be reordered without touching a single report formula.

Private Const SHEET_SUM As String = "Summary"
Private Const SHEET_REPORTS As String = "Reports"
Private Const DATA_SPILL As String = "_Data!$A$10#"
Private Const DATA_HDR As String = "_Data!$A$9:$AZ$9"
Private Const EXPORT_SIG_CELL As String = "AN1"
Private Const EXPORT_WHEN_CELL As String = "AN2"

' A column of the consolidated range, found by its header text.
Private Function C(ByVal Header As String) As String
    C = "INDEX(" & DATA_SPILL & ",,MATCH(""" & Header & """," & DATA_HDR & ",0))"
End Function

' SUMIFS of one column, grouped by the Location/Printer/Paper stock key. The
' criteria are arrays, which is what makes the result spill down alongside the
' keys instead of needing a formula per row.
Private Function SumBy(ByVal Header As String) As String
    SumBy = "SUMIFS(" & C(Header) & "," & C("Location") & ",lo," & C("Printer") & ",pr," & C("Paper Stock") & ",st)"
End Function

Public Sub BuildReportSheets()
    BuildSummary
    BuildReports
End Sub

' ============================================================== Summary ===
Public Sub BuildSummary()
    Dim ws As Worksheet, f As String

    Set ws = SheetNamed(SHEET_SUM, 1)
    UnlockSheet ws
    ws.Cells.Clear
    ClearSpillArea ws

    ws.Range("A1").Value = "Summary"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Formula = "=SET_ORG&"" - ""&SET_DEPT&""  |  ""&SET_SITE_NAME"
    ws.Range("A3").Value = "Every print job recorded in this workbook, by print room, printer and paper stock. " & _
        "Figures are live - they update as jobs are added and need no refreshing."

    ' --- totals, above the table ------------------------------------------
    ' Above, not below: the detail spills to an unknown height, so anything
    ' placed under it would be overwritten the moment a job is added.
    ws.Range("A5").Value = "Totals"
    ws.Range("A5").Font.Bold = True
    TotalCell ws, "B6", "Jobs", "=IFERROR(COUNTA(" & C("Job ID") & "),0)"
    TotalCell ws, "D6", "Gross", "=IFERROR(SUM(" & C("Gross Cost") & "),0)"
    TotalCell ws, "F6", "Disregarded", "=IFERROR(SUM(" & C("Disregarded") & "),0)"
    TotalCell ws, "H6", "Chargeable", "=IFERROR(SUM(" & C("Chargeable Cost") & "),0)"
    ' Snag list item 1c: the chargeable total split by paid status. Paid is
    ' "total minus paid" rather than a separate <>"Yes" SUMIFS, so the two
    ' always reconcile exactly to Chargeable by construction - a blank Paid
    ' (a job that predates the column) falls into Unpaid either way.
    TotalCell ws, "J6", "Paid", "=IFERROR(SUMIFS(" & C("Chargeable Cost") & "," & C("Paid") & ",""Yes""),0)"
    TotalCell ws, "L6", "Unpaid", "=IFERROR(SUM(" & C("Chargeable Cost") & ")-SUMIFS(" & C("Chargeable Cost") & "," & C("Paid") & ",""Yes""),0)"

    ' --- headers -----------------------------------------------------------
    WriteHeaderRow ws, 9, Array("Location", "Printer", "Paper stock", "Type", "Unit", _
                                "Jobs", "Qty", "Area m2", "Paper cost", _
                                "Consumable cost", "Gross", "Disregarded", "Chargeable")

    ' --- the one formula -----------------------------------------------------
    ' Three-column key - Location, Printer, Paper Stock - rather than the
    ' two-column Location x Paper Stock key this replaces. Papers common to
    ' more than one printer model are deliberately not combined: two printers
    ' sharing a stock still get their own row each, because the cost per
    ' print differs by printer (consumable rate) even when the paper does not.
    f = "=IFERROR(LET(" & _
        "keys,SORT(UNIQUE(HSTACK(" & C("Location") & "," & C("Printer") & "," & C("Paper Stock") & "))),"
    f = f & "lo,INDEX(keys,,1),pr,INDEX(keys,,2),st,INDEX(keys,,3),"
    ' The two built-in Supplied stocks are not tblPapers rows (modCatalog.
    ' AddBuiltInStocks), so each lookup falls back to their fixed values
    ' before giving up with "(not in Papers)".
    Dim isRoll As String, isSheet As String
    isRoll = "st=""" & SUPPLIED_ROLL & """"
    isSheet = "st=""" & SUPPLIED_SHEET & """"
    f = f & "HSTACK(lo,pr,st," & _
        "IFNA(XLOOKUP(st,tblPapers[Description],tblPapers[Paper type]),IF(OR(" & isRoll & "," & isSheet & "),""Student supplied"",""(not in Papers)"")),"
    f = f & "IFNA(IF(XLOOKUP(st,tblPapers[Description],tblPapers[Measure])=""Sheet"",""sheets"",""metres""),IF(" & isRoll & ",""metres"",IF(" & isSheet & ",""sheets"",""-""))),"
    f = f & "COUNTIFS(" & C("Location") & ",lo," & C("Printer") & ",pr," & C("Paper Stock") & ",st),"
    f = f & SumBy("Qty") & "," & SumBy("Area m2") & "," & SumBy("Paper Cost") & ","
    f = f & SumBy("Consumable Cost") & "," & SumBy("Gross Cost") & ","
    f = f & SumBy("Disregarded") & "," & SumBy("Chargeable Cost") & ")),"
    f = f & """No print jobs have been recorded yet."")"
    ws.Range("A10").Formula2 = f

    FormatSummary ws
    FormatSummaryErrors ws
    DrawLegend ws
    RelockSheet ws
End Sub

Private Sub FormatSummary(ByVal ws As Worksheet)
    ' Warm tint (0.9.14), matching the Reports page's own totals header (row
    ' 12 there) rather than its blue results-table header - this table is
    ' itself a totals breakdown (by location/printer/paper stock), the same
    ' category as Reports' "Matching" row, not a per-job record list.
    ws.Range("A9:M9").Interior.Color = RGB(244, 232, 222)
    ws.Range("G10:H2000").NumberFormat = "#,##0.00"
    ws.Range("I10:M2000").NumberFormat = CurrencyFormatCode()
    ws.Range("D6").NumberFormat = CurrencyFormatCode()
    ws.Range("F6").NumberFormat = CurrencyFormatCode()
    ws.Range("H6").NumberFormat = CurrencyFormatCode()
    ws.Range("J6").NumberFormat = CurrencyFormatCode()
    ws.Range("L6").NumberFormat = CurrencyFormatCode()
    ws.Columns("A:M").ColumnWidth = 14
    ws.Columns("A:C").ColumnWidth = 24
    ws.Rows(9).Font.Bold = True
End Sub

' Phase 8's "error state" (design doc §11): the one place the workbook's own
' formulas already flag a genuine data-integrity break, as opposed to a
' routine per-row validation nag - a paper stock a job was costed against has
' since been renamed or removed from tblPapers, so the Type lookup in
' A10's spilled formula falls back to the literal "(not in Papers)" (IFNA in
' BuildSummary above). Like Status, this is pure spilled-formula output with
' no per-cell VBA hook, so conditional formatting is the only way to colour
' it. Red reuses the exact colour the Reports disjoint-criteria warning and
' the "needs export" flag already use (modReports.BuildReports, modExport.
' RefreshExportStatus), so "error" reads the same everywhere it appears.
Private Sub FormatSummaryErrors(ByVal ws As Worksheet)
    Dim rng As Range, fc As FormatCondition
    Set rng = ws.Range("D10:D2000")
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=D10=""(not in Papers)""")
    fc.Font.Color = RGB(176, 0, 32)
    fc.Font.Bold = True
End Sub

' Phase 8's Summary legend (design doc §11 / architecture §16.2): explains
' the four everyday cell-role colours - not the exceptional Warning/Error
' states above, which explain themselves by firing, and not the niche
' Snapshot/historical role, which lives collapsed in a column group nobody
' opens day to day. Sits at O9 down: clear of the report table (A:M) and
' clear of the buttons InitialiseWorkbook draws at column O rows 1/3/5/7, and
' level with the table's own header row so it reads as "this explains that".
' Colours are the exact ones already in use elsewhere on this sheet/Reports,
' so the swatches actually match what they claim to explain.
Private Sub DrawLegend(ByVal ws As Worksheet)
    ws.Range("O9").Value = "Cell colours"
    ws.Range("O9").Font.Bold = True

    LegendRow ws, 10, "Type your own values here", RGB(255, 255, 255), False, False, True
    LegendRow ws, 11, "Calculated automatically", RGB(217, 217, 217), True, False, False
    LegendRow ws, 12, "Workbook configuration", RGB(222, 232, 244), False, False, False
    LegendRow ws, 13, "Reference information (read-only)", RGB(217, 217, 217), False, True, False

    ws.Columns("O").ColumnWidth = 4
    ws.Columns("P").ColumnWidth = 30
End Sub

Private Sub LegendRow(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal Label As String, _
                       ByVal Fill As Long, ByVal Italic As Boolean, ByVal GreyText As Boolean, _
                       ByVal BlueBorder As Boolean)
    Dim swatch As Range
    Set swatch = ws.Cells(RowNo, 15) ' column O
    swatch.Interior.Color = Fill
    If BlueBorder Then
        With swatch.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(46, 100, 168)
        End With
    End If

    With ws.Cells(RowNo, 16) ' column P
        .Value = Label
        .Font.Italic = Italic
        If GreyText Then .Font.Color = RGB(110, 110, 110)
    End With
End Sub

' ============================================================= Reports ===
' The criteria collapse to TRUE when blank, which is what gives the multi-
' case behaviour AT-12 tests without a single branch in the formula.
Private Function Criteria() As String
    Dim s As String

    ' The leading term exists to fix the SHAPE of the result, and without it
    ' the whole sheet breaks in its most ordinary state. A blank criterion
    ' makes IF(...) return a scalar TRUE, so with all four boxes empty the
    ' product is the scalar 1 - and FILTER(column, 1) is #CALC!, because the
    ' include argument has to be as tall as the data. Job ID is never blank in
    ' the consolidated range (7.1 filters those rows out), so this is an array
    ' of TRUE the right height, and every criterion below multiplies into it.
    s = "(" & C("Job ID") & "<>"""")"

    s = s & "*IF($B$4="""",TRUE,ISNUMBER(SEARCH($B$4," & C("Student Name") & ")))"
    s = s & "*IF($B$5="""",TRUE,TRIM(" & C("Student No") & "&"""")=TRIM($B$5&""""))"

    ' Dates are coerced with *1 on BOTH sides. A user typing 16/09/2026 into
    ' an unformatted cell can leave text there, and number >= text compares as
    ' FALSE for every row without raising anything - a filter that silently
    ' returns nothing. Text that will not coerce is treated as no filter at
    ' all rather than as a date of zero.
    '
    ' The column holds date AND time, so the end test is < end+1: otherwise a
    ' job logged at 16:30 on the closing date falls outside its own range.
    '
    ' $B$7/$B$8, not $B$6/$B$7: row 6 is deliberately blank (2026-09-26) - a
    ' gap between the Student/Department pair above and the date pair here,
    ' the same row the date boxes already sat at before that gap was added.
    s = s & "*IF($B$7="""",TRUE,IF(ISERROR($B$7*1),TRUE," & _
        "IFERROR(" & C("Date/Time") & "*1,0)>=$B$7*1))"
    s = s & "*IF($B$8="""",TRUE,IF(ISERROR($B$8*1),TRUE," & _
        "IFERROR(" & C("Date/Time") & "*1,0)<$B$8*1+1))"

    ' Technician, Printer and Paper Stock are dropdowns (RefreshReportFilterLists),
    ' not free text, so an exact match is what "choose one from the list"
    ' means - unlike Student name, there is no fragment to search for.
    ' Rows shifted down one (2026-09-27, see the CritCell call sites) to make
    ' room for Location at $F$4, immediately below.
    s = s & "*IF($F$5="""",TRUE," & C("Technician") & "=$F$5)"
    s = s & "*IF($F$6="""",TRUE," & C("Printer") & "=$F$6)"
    s = s & "*IF($F$7="""",TRUE," & C("Paper Stock") & "=$F$7)"

    ' Quantity: exact match, blank ignored, *1-coerced the same way the date
    ' boxes are so a value left as text by an unformatted cell is treated as
    ' no filter rather than as a quantity of zero.
    s = s & "*IF($F$8="""",TRUE,IF(ISERROR($F$8*1),TRUE," & _
        "IFERROR(" & C("Qty") & "*1,0)=$F$8*1))"

    ' Location (print room), added 2026-09-25 - a dropdown of registered print
    ' rooms (RefreshReportFilterLists, AllLocationCodes), same exact-match
    ' treatment as Technician/Printer/Paper Stock above. Moved from $O$4 to
    ' $F$4 (2026-09-27, see its CritCell call site) once Technician/Printer/
    ' Paper Stock/Quantity shifted down a row and freed it.
    s = s & "*IF($F$4="""",TRUE," & C("Location") & "=$F$4)"

    Criteria = s
End Function

' The sortable columns, in results-table order (Job ID excluded - it is a
' hidden correlation column, not something offered in the sort-by list).
' Student name/no and Paid added 2026-09-22 (snag list items 2a, 2d).
Private Function ResultHeaders() As Variant
    ResultHeaders = Array("Date/Time", "Location", "Student name", "Student no", _
                          "Printer", "Paper stock", _
                          "Qty", "Unit", "Area m2", "Paper cost", _
                          "Consumable cost", "Gross", "Disregarded", "Chargeable", "Paid", _
                          "Technician", "Notes")
End Function

Private Function QuotedList(ByVal items As Variant) As String
    Dim i As Long, s As String
    For i = LBound(items) To UBound(items)
        If Len(s) > 0 Then s = s & ","
        s = s & """" & items(i) & """"
    Next i
    QuotedList = s
End Function

Public Sub BuildReports()
    Dim ws As Worksheet, f As String, ok As String, hdrs As Variant

    Set ws = SheetNamed(SHEET_REPORTS, 2)
    UnlockSheet ws

    ' The Export report stamp (Problem 2's refinement) has to survive this
    ' rebuild the same way tblLocations' Last export/Export sig survive
    ' RefreshLocations' registry rebuild - captured before Cells.Clear, put
    ' back once the sheet exists again.
    Dim savedSig As String, savedWhen As Variant
    savedSig = CStr(ws.Range(EXPORT_SIG_CELL).Value)
    savedWhen = ws.Range(EXPORT_WHEN_CELL).Value

    ws.Cells.Clear
    ok = Criteria
    hdrs = ResultHeaders

    ws.Range("A1").Value = "Reports"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Find and filter print jobs across every room in this workbook. Results update as you type - " & _
        "there is no search button. Leave a box empty to ignore it."

    ' Label | input | hint | gap, in that order (snag list item 2) - the
    ' input sits immediately right of its label, and the unused column at the
    ' end of each group is what separates it from the next.
    ' "Student/Department" labels (snag 4b, 2026-09-22): Student Name/No is
    ' also used free-text for department charging (D19) - the underlying
    ' column headers stay "Student Name"/"Student No" (renaming those would
    ' break every header-name lookup that already reads them, for no real
    ' capability gained), but the on-screen wording someone actually reads
    ' here reads sensibly either way.
    CritCell ws, "A4", "B4", "Student/Department name"
    CritCell ws, "A5", "B5", "Student/Department number"
    ' Row 6 left blank (2026-09-26) - a gap between the Student/Department
    ' pair above and the date pair below, freed up by trimming the two-row
    ' gap under the title (rows 3-4) to one (row 3 only) rather than growing
    ' the sheet by a row overall.
    CritCell ws, "A7", "B7", "From date"
    CritCell ws, "A8", "B8", "To date"
    ws.Range("B7:B8").NumberFormat = "dd/mm/yyyy"
    AddDateValidation ws.Range("B7"), "From date", "Leave blank for no start date."
    AddDateValidation ws.Range("B8"), "To date", "Jobs logged at any time on this date are included."

    ' Labels at D, not E (2026-09-26 fix): E is "Printer" in the results
    ' table, one of the columns ApplyReportsMinimumColumns hides by header
    ' name - and that Hidden reaches every row on the sheet, not just the
    ' results table, so a label parked at E4:E8 rendered invisible under the
    ' default view even though its own input cell (F, "Paper stock", always
    ' kept) showed fine with nothing beside it. D ("Student no") is one of
    ' the kept columns and unused on these rows - previously the blank
    ' spacer between this group and the Student/Department one at A:B, now
    ' doing double duty as the label column instead. The one-column visual
    ' gap between groups is gone, but a readable label beats a tidy gap to a
    ' label nobody could see.
    '
    ' Technician/Printer/Paper stock/Quantity moved down one row, D5:D8/F5:F8
    ' (2026-09-27, direct user request), freeing D4/F4 for the Location
    ' (print room) filter immediately below - see that CritCell call for why
    ' it moved out of N4/O4.
    CritCell ws, "D5", "F5", "Technician"
    CritCell ws, "D6", "F6", "Printer"
    CritCell ws, "D7", "F7", "Paper stock"
    CritCell ws, "D8", "F8", "Quantity"

    ' Location (print room) filter. Originally parked at N4/O4 (2026-09-25) to
    ' dodge E/G/H/I:M, the columns ApplyReportsMinimumColumns hides entirely
    ' by results-header name (the same trap O10's own 2026-09-22 comment
    ' names) - a label or input parked there can render hidden or orphaned
    ' under the default view. Moved to D4/F4 (2026-09-27, direct user
    ' request) - the row the Technician/Printer/Paper Stock/Quantity group
    ' vacated by shifting down one row (above) - rather than staying at N4/O4,
    ' so the whole filter block reads top-to-bottom as one contiguous group:
    ' Location, Technician, Printer, Paper stock, Quantity.
    CritCell ws, "D4", "F4", "Location (print room)"

    ' Sort by/direction sit below the filters, above the totals row (snag list
    ' item 3) rather than beside the Technician/Printer/Paper/Quantity group -
    ' row 9 is the "name and number don't match" warning below, so this is the
    ' one free row between the filters and Matching.
    CritCell ws, "A10", "B10", "Sort by"
    CritCell ws, "D10", "F10", "Sort direction"
    AddList ws.Range("B10"), QuotedList(hdrs), "Sort by", "Which column to sort the results by."
    AddList ws.Range("F10"), """Ascending"",""Descending""", "Sort direction", "Which way to sort."

    ' Snag list item 2a: student name/number are not shown or exported
    ' unless switched on - defaults to No (data protection: opt in to
    ' reveal, not opt out). This is a display option, not a filter criterion,
    ' so it moves onto the sort-controls row rather than sharing row 9 with
    ' the (unrelated) name/number warning, where it used to sit at E9/F9.
    '
    ' N10/O10, not the next free columns (I/J) after Sort direction: this
    ' sheet's minimum-columns view (ApplyReportsMinimumColumns) hides entire
    ' COLUMNS by results-header name, and that hide reaches every row on the
    ' sheet, not just the results table - which is exactly why E9's label was
    ' invisible before (E is "Printer", one of the hidden columns) while its
    ' F9 dropdown (F is "Paper stock", always kept visible) still showed with
    ' nothing beside it. N and O are "Chargeable"/"Paid" - both permanently
    ' kept - and free on this row, so label and dropdown survive the default
    ' view together. (Sort direction's own label has this same latent problem
    ' and was fixed the same way, 2026-09-26 - see its own CritCell call
    ' above, D10 rather than E10.)
    CritCell ws, "N10", "O10", "Show names"
    AddList ws.Range("O10"), """Yes"",""No""", "Show names", "Yes shows the student/department name and number in the results and any export. No (the default) blanks them, for data protection."
    If Len(Trim$(CStr(ws.Range("O10").Value))) = 0 Then ws.Range("O10").Value = "No"

    ' Spec 14.1: both criteria given, neither matching the other.
    ws.Range("A9").Formula2 = "=IF(OR($B$4="""",$B$5=""""),""""," & _
        "IF(IFERROR(ROWS(FILTER(" & C("Job ID") & "," & ok & ")),0)=0," & _
        """That name and that number do not appear together on any record - check both."",""""))"
    ws.Range("A9").Font.Color = RGB(176, 0, 32)
    ws.Range("A9").Font.Bold = True

    ' --- totals for the current selection ---------------------------------
    ' Row 11 is left blank (snag list item 1) - a gap between the sort
    ' controls and Matching, matching the gap that already separates the
    ' totals from the results header below.
    '
    ' Snag list item 2d hides several results columns by default (§8.3) -
    ' hiding a column hides the WHOLE column, every row, so a totals cell
    ' sharing a column with a hidden results column would display blank,
    ' the same class of bug §4.1's 2026-09-22 fix found on the location
    ' sheets. Rather than relocate the whole "Matching" block, each metric's
    ' label now sits directly ABOVE its value (row 12/13, same column)
    ' instead of to its left (MatchTotal, below), and both land only on
    ' columns the minimum-columns view never hides: B, C, D, F, N, O -
    ' Location/Student Name/Student No/Paper Stock/Chargeable Cost/Paid's
    ' own columns, reused here purely because they're guaranteed visible,
    ' not because a total means the same thing as whatever heads that
    ' column further down the sheet. When the hidden columns between them
    ' collapse (the default state), B/C/D/F/N/O end up rendering adjacent
    ' anyway, so nothing looks gapped in the common case.
    ws.Range("A12").Value = "Matching"
    ws.Range("A12").Font.Bold = True
    MatchTotal ws, "B", "Jobs", "=IFERROR(ROWS(FILTER(" & C("Job ID") & "," & ok & ")),0)"
    MatchTotal ws, "C", "Gross", "=IFERROR(SUM(FILTER(" & C("Gross Cost") & "," & ok & ")),0)"
    MatchTotal ws, "D", "Disregarded", "=IFERROR(SUM(FILTER(" & C("Disregarded") & "," & ok & ")),0)"
    MatchTotal ws, "F", "Chargeable", "=IFERROR(SUM(FILTER(" & C("Chargeable Cost") & "," & ok & ")),0)"
    ' Snag list item 1c: the matching chargeable total split by paid status,
    ' same "total minus paid" reconciliation as Summary's J6/L6 - a blank
    ' Paid (a job that predates the column) falls into Unpaid either way.
    ' FILTER's third (if_empty) argument matters here specifically: with
    ' nothing yet marked Paid, paidOk matches zero rows, and a plain
    ' FILTER(array, all-false) raises #CALC! rather than returning an empty
    ' array - which would otherwise poison the whole subtraction below (an
    ' error on EITHER side of a "-" makes the whole expression an error) and
    ' get masked by the outer IFERROR into a false "0", not the real total.
    ' Found by testing the everyday "nothing paid yet" case, not just the
    ' "one row marked Yes" case test-paid.ps1 already covered.
    Dim paidOk As String
    paidOk = ok & "*(" & C("Paid") & "=""Yes"")"
    MatchTotal ws, "N", "Paid", "=IFERROR(SUM(FILTER(" & C("Chargeable Cost") & "," & paidOk & ",0)),0)"
    MatchTotal ws, "O", "Unpaid", "=IFERROR(SUM(FILTER(" & C("Chargeable Cost") & "," & ok & ",0))-SUM(FILTER(" & C("Chargeable Cost") & "," & paidOk & ",0)),0)"
    ws.Range("C13").NumberFormat = CurrencyFormatCode()
    ws.Range("D13").NumberFormat = CurrencyFormatCode()
    ws.Range("F13").NumberFormat = CurrencyFormatCode()
    ws.Range("N13").NumberFormat = CurrencyFormatCode()
    ws.Range("O13").NumberFormat = CurrencyFormatCode()

    ' --- the records ------------------------------------------------------
    ' Job ID is appended after Notes and hidden - the correlation key that
    ' maps a visible row back to its source location sheet and table row for
    ' the Reports-page delete. Nothing else moves, so existing column
    ' positions are untouched.
    WriteHeaderRow ws, 15, hdrs
    ws.Cells(15, UBound(hdrs) - LBound(hdrs) + 2).Value = "Job ID"

    ' Sorting: the whole FILTER result is bound to res once via LET, then
    ' re-ordered by whichever column K5 names, found by matching its header
    ' text against the same header list the table itself uses so the two
    ' can never drift apart. IFERROR around the SORTBY step is what keeps
    ' FILTER's own "no rows matched"/"no data at all" fallback text intact
    ' when there is nothing to sort - SORTBY on that text would otherwise
    ' error and the outer IFERROR would show the wrong one of the two
    ' messages.
    ' Student Name/No (snag 2a): the toggle at O10 blanks the VALUES, not
    ' just the column - satisfies data protection even if someone unhides
    ' the column, since there is nothing behind it to reveal. ExportReport
    ' Snapshot copies these live values as-is, so the toggle is respected
    ' by the export automatically with no separate export-side logic.
    '
    ' Built with Chr(34) for the formula's own empty-string literal rather
    ' than hand-counting doubled quotes in a VBA string literal - a single
    ' wrong quote count here is a silent formula-text bug, not a compile
    ' error, so it is worth avoiding the manual counting entirely.
    '
    ' IF($O$10="Yes", <column array>, "") is NOT the same as an elementwise
    ' per-row blank - the CONDITION here is a bare scalar (one toggle cell),
    ' so with O10="No" the whole IF collapses to the single scalar "" rather
    ' than a column of blanks the same height as every other HSTACK
    ' argument. HSTACK does not broadcast a scalar against a column
    ' (modRegistry's own §7.1 comment already names this trap for the
    ' consolidated-range formula) - the visible symptom here was rows
    ' beyond the first silently showing FILTER's "no data" fallback text
    ' instead of real job data. Fixed by nesting the toggle INSIDE an outer
    ' IF whose own condition (Job ID <> "") is already a real per-row array
    ' - once the outer IF is evaluating elementwise, the inner one is too,
    ' so $O$10="Yes" correctly re-tests against the SAME scalar for every
    ' row while Student Name/No resolve per row as normal.
    Dim q As String, studentOn As String, rowShape As String, sName As String, sNo As String
    q = Chr(34)
    studentOn = "$O$10=" & q & "Yes" & q
    rowShape = C("Job ID") & "<>" & q & q
    sName = "IF(" & rowShape & ",IF(" & studentOn & "," & C("Student Name") & "," & q & q & ")," & q & q & ")"
    sNo = "IF(" & rowShape & ",IF(" & studentOn & "," & C("Student No") & "," & q & q & ")," & q & q & ")"

    ' A row that predates the Paid column (§5, blank Paid counts as unpaid)
    ' reads back from INDEX as the NUMBER 0, not an empty string - the
    ' classic "reference to a genuinely blank cell returns 0" Excel
    ' behaviour - so the results table would otherwise display a literal
    ' "0" for those rows. &"" does NOT fix this (0&"" is still the text
    ' "0", just stringified) - Paid only ever legitimately holds "Yes",
    ' "No" or blank, never a real 0, so testing for =0 specifically catches
    ' exactly the phantom-zero case and nothing else.
    Dim paidText As String
    paidText = "IF(" & C("Paid") & "=0," & q & q & "," & C("Paid") & ")"

    f = "=IFERROR(LET(" & _
        "res,FILTER(HSTACK(" & C("Date/Time") & "," & C("Location") & "," & sName & "," & sNo & _
        "," & C("Printer") & "," & C("Paper Stock") & "," & C("Qty") & "," & C("Unit") & "," & C("Area m2") & _
        "," & C("Paper Cost") & "," & C("Consumable Cost") & "," & C("Gross Cost") & _
        "," & C("Disregarded") & "," & C("Chargeable Cost") & "," & paidText & "," & C("Technician") & _
        "," & C("Notes") & "," & C("Job ID") & ")," & ok & ",""No print jobs match those criteria.""),"
    f = f & "hdrs,{" & QuotedList(hdrs) & "},"
    f = f & "sortIdx,IFERROR(MATCH($B$10,hdrs,0),0),"
    f = f & "dir,IF($F$10=""Descending"",-1,1),"
    f = f & "IF(sortIdx=0,res,IFERROR(SORTBY(res,INDEX(res,,sortIdx),dir),res))"
    f = f & "),""No print jobs have been recorded yet."")"
    ws.Range("A16").Formula2 = f
    ws.Columns(UBound(hdrs) - LBound(hdrs) + 2).Hidden = True

    ws.Range(EXPORT_SIG_CELL).Value = savedSig
    If IsDate(savedWhen) Then
        ws.Range(EXPORT_WHEN_CELL).Value = savedWhen
        ws.Range(EXPORT_WHEN_CELL).NumberFormat = "dd/mm/yyyy hh:mm"
    End If
    ws.Range(EXPORT_SIG_CELL).EntireColumn.Hidden = True

    BuildBreakdowns ws, ok
    FormatReports ws

    ' The filter labels/hints (CritCell, above) were written and wrapped
    ' before this point, while columns A/C/D/N still sat at Excel's
    ' factory-default width - FormatReports's ColumnWidth calls, just above,
    ' are the first thing to widen them. Re-fitting now, with the real widths
    ' in place, is what keeps rows 4:10 sized for the text that actually
    ' fits per line rather than for the cramped default - see CritCell's own
    ' comment for how this was found.
    ws.Rows("4:10").AutoFit

    ' Snag list item 2d: the results table keeps every column, but only the
    ' documented minimum stays visible by default - the rest are hidden
    ' (never removed), reusing the exact same technique as the location
    ' sheets' reduced-clutter view (modInit.ApplyColumnVisibility) even
    ' though this table isn't a ListObject, so that function can't be
    ' called directly. Not a user-facing toggle here, per the snag list's
    ' own note that customising which columns show is a future enhancement.
    ApplyReportsMinimumColumns ws, hdrs

    ' Last, deliberately. ApplyTo (called from RefreshReportFilterLists) does
    ' its own Unlock/RelockSheet on ws as a self-contained operation - called
    ' any earlier, its RelockSheet would re-protect the sheet mid-build and
    ' every Validation.Add after it (the sort dropdowns above) would fail
    ' with a bare 1004 on the now-protected cells.
    RefreshReportFilterLists ws

    ' Row group (snag list item 15): the whole filter block - both criteria
    ' groups, the name/number warning and the sort controls - collapses
    ' together, so a user who has already set filters can hide the controls
    ' without losing them. Grouping needs the sheet unprotected, same reason
    ' RefreshReportFilterLists' own ApplyTo calls do their own Unlock/Relock.
    UnlockSheet ws
    ws.Rows("4:10").Group
    RelockSheet ws
End Sub

' Breakdowns sit to the RIGHT of the record list, not beneath it. The list
' spills to a height nobody can predict, so anything below it would be
' displaced the moment one more job matched.
'
' Shifted from Q/U to T/X on 2026-09-22 (snag 2a/2d): the results table grew
' from 14 to 17 columns (Student Name/No added, Paid added), so the old Q15
' start now sits inside the table itself (Notes lands on Q). The +3 shift
' preserves the exact same relative gaps this had before (one blank column
' after the hidden Job ID column, then straight into "By print room").
Private Sub BuildBreakdowns(ByVal ws As Worksheet, ByVal ok As String)
    ws.Range("T15").Value = "By print room"
    ws.Range("T15").Font.Bold = True
    ws.Range("T16").Formula2 = GroupFormula(ok, "Location")

    ws.Range("X15").Value = "By paper stock"
    ws.Range("X15").Font.Bold = True
    ws.Range("X16").Formula2 = GroupFormula(ok, "Paper Stock")

    ' Each block spills as key | Jobs | Gross | Chargeable, so the money
    ' columns are the third and fourth - Jobs is a count and must not be
    ' formatted as currency.
    ws.Range("U17:U2000").NumberFormat = "#,##0"
    ws.Range("V17:W2000").NumberFormat = CurrencyFormatCode()
    ws.Range("Y17:Y2000").NumberFormat = "#,##0"
    ws.Range("Z17:AA2000").NumberFormat = CurrencyFormatCode()
End Sub

' Group the filtered records by one column. SUMIFS cannot be used here: its
' arguments must be ranges, and once FILTER has been applied these are arrays.
' BYROW with a LAMBDA does the per-key aggregation instead.
Private Function GroupFormula(ByVal ok As String, ByVal KeyHeader As String) As String
    Dim s As String
    ' The criteria are bound once as `ok` rather than pasted into each FILTER.
    ' Three copies of a 700-character expression is not just wasteful - it is
    ' three places for the four criteria to drift apart.
    s = "=IFERROR(LET(" & _
        "ok," & ok & "," & _
        "k,FILTER(" & C(KeyHeader) & ",ok)," & _
        "g,FILTER(" & C("Gross Cost") & ",ok)," & _
        "c,FILTER(" & C("Chargeable Cost") & ",ok)," & _
        "u,SORT(UNIQUE(k))," & _
        "VSTACK(HSTACK(""" & KeyHeader & """,""Jobs"",""Gross"",""Chargeable"")," & _
        "HSTACK(u," & _
        "BYROW(u,LAMBDA(x,SUM(--(k=x))))," & _
        "BYROW(u,LAMBDA(x,SUM(FILTER(g,k=x,0))))," & _
        "BYROW(u,LAMBDA(x,SUM(FILTER(c,k=x,0))))))" & _
        "),"""")"
    GroupFormula = s
End Function

' Column letters below reflect the 2026-09-22 layout: A Date/Time, B
' Location, C Student Name, D Student No, E Printer, F Paper Stock, G
' Qty, H Unit, I Area m2, J Paper Cost, K Consumable Cost, L Gross
' Cost, M Disregarded, N Chargeable Cost, O Paid, P Technician, Q Notes,
' (R Job ID, hidden).
Private Sub FormatReports(ByVal ws As Worksheet)
    ' Matching's header row (row 12) gets the same subtle-tint treatment as
    ' the results header below, in a warmer tone so the two bands read as
    ' related but distinct - RGB(244, 232, 222) is RGB(222, 232, 244)'s own
    ' red/blue channels swapped, keeping the identical lightness/saturation.
    ws.Range("A12:O12").Interior.Color = RGB(244, 232, 222)
    ws.Range("A15:Q15").Interior.Color = RGB(222, 232, 244)
    ws.Range("A16:A2000").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range("G16:I2000").NumberFormat = "#,##0.00"
    ws.Range("J16:N2000").NumberFormat = CurrencyFormatCode()
    ws.Columns("A:Q").ColumnWidth = 14
    ws.Columns("B:F").ColumnWidth = 22
    ws.Columns("Q").ColumnWidth = 30
    ws.Columns("A").ColumnWidth = ColWidthForPx(180)  ' Date/Time, 180px
    ws.Columns("T").ColumnWidth = 22
    ws.Columns("X").ColumnWidth = 22
    ws.Rows(15).Font.Bold = True
End Sub

' ======================================================== delete visible ===
' Problem 3: bulk pruning, deliberately confined to the Reports page and
' operating on whatever is currently visible under the active filters.
' modJobs.RemoveRow (single row) and ClearAll (whole sheet) on individual
' location sheets are sufficient and stay as they are - this is the
' higher-friction, encourages-a-backup-first path for pruning across rooms.
Public Sub DeleteVisibleReports()
    Dim ws As Worksheet, rng As Range, n As Long
    Dim locCol As Long, jobCol As Long, i As Long, loc As String
    Dim rooms As clsDict, k As Variant, breakdown As String, staleWarn As String
    Dim curSig As String, lastSig As String

    Set ws = ReportsSheet()
    If ws Is Nothing Then
        Say "The Reports sheet could not be found.", "Run Refresh Locations first."
        Exit Sub
    End If

    On Error Resume Next
    Set rng = ws.Range("A16").SpillingToRange
    On Error GoTo 0
    If rng Is Nothing Then
        Say "There is nothing to delete.", "The Reports sheet has no results under the current filters."
        Exit Sub
    End If
    ' A16 may be spilling FILTER's own "no jobs match"/"no jobs recorded"
    ' fallback TEXT rather than real rows - see FilteredSig's own comment.
    If Not IsNumeric(rng.Cells(1, 1).Value2) Then
        Say "There is nothing to delete.", CStr(rng.Cells(1, 1).Value)
        Exit Sub
    End If
    n = rng.Rows.Count

    locCol = ColByHeader(ws, 15, "Location")
    jobCol = ColByHeader(ws, 15, "Job ID")
    If locCol = 0 Or jobCol = 0 Then
        Say "The Reports sheet layout looks wrong.", "The Location or Job ID column could not be found.", "Rebuild the report sheets (Refresh Locations), then try again."
        Exit Sub
    End If

    ' Broken down by room, following the same "show what you're about to
    ' lose" pattern as modJobs.RemoveRow/ClearAll.
    Set rooms = New clsDict
    For i = 1 To n
        loc = CStr(rng.Cells(i, locCol).Value)
        If rooms.Exists(loc) Then
            rooms.Add loc, CLng(rooms.Item(loc)) + 1
        Else
            rooms.Add loc, 1
        End If
    Next i
    For Each k In rooms.Keys
        breakdown = breakdown & "  " & CStr(k) & ": " & rooms.Item(CStr(k)) & vbCrLf
    Next k

    ' Warn when the current filtered set does not match what the last
    ' Export report run actually captured - never exported, or the filters
    ' or underlying data have changed since. Replaces the original "warn if
    ' not exported to CSV" wording, since deletion here is keyed off the
    ' Export report signature rather than per-location CSV export status.
    curSig = FilteredSig(ws)
    lastSig = ReportsExportSig(ws)
    If Len(lastSig) = 0 Then
        staleWarn = vbCrLf & vbCrLf & "Export report has never been run for a filtered set like this one - " & _
            "there is no up-to-date report or backup of what is about to be deleted."
    ElseIf StrComp(lastSig, curSig, vbBinaryCompare) <> 0 Then
        staleWarn = vbCrLf & vbCrLf & "The filters or underlying data have changed since the last Export report - " & _
            "there is no up-to-date report or backup of what is about to be deleted."
    End If

    If Not Ask("Delete " & n & " visible record" & IIf(n = 1, "", "s") & " from the Reports page?" & vbCrLf & vbCrLf & _
        breakdown & staleWarn & vbCrLf & vbCrLf & "This cannot be undone.", "Delete visible records") Then Exit Sub

    DeleteVisibleReportsConfirmed ws, rng, locCol, jobCol, n
End Sub

' The actual deletion. Public so a test can call it directly, bypassing the
' Ask() gate above - Ask() always declines under SetQuiet, same reason
' modJobs.RemoveRow/ClearAll and modImport.ApplyImportConfirmed are split
' this way.
Public Sub DeleteVisibleReportsConfirmed(ByVal ws As Worksheet, ByVal rng As Range, _
                                         ByVal LocCol As Long, ByVal JobCol As Long, ByVal n As Long)
    Dim i As Long, targetWs As Worksheet, lo As ListObject, rowIdx As Long
    Dim deleted As Long, missing As Long, detail As String
    Dim jobIds() As String, locs() As String

    On Error GoTo Fail

    ' Copied into memory FIRST. rng points at a live spilled formula range -
    ' the moment the first row is deleted from a source table, _Data (and
    ' this sheet's own FILTER) shrink, and rng's cells would be reading a
    ' moving target for every row after the first.
    ReDim jobIds(1 To n)
    ReDim locs(1 To n)
    For i = 1 To n
        jobIds(i) = CStr(rng.Cells(i, JobCol).Value)
        locs(i) = CStr(rng.Cells(i, LocCol).Value)
    Next i

    AppOff
    For i = 1 To n
        Set targetWs = SheetForCode(locs(i))
        rowIdx = 0
        If Not targetWs Is Nothing Then
            Set lo = JobsTable(targetWs)
            If Not lo Is Nothing Then rowIdx = FindReportRow(lo, jobIds(i))
        End If
        If rowIdx > 0 Then
            UnlockSheet targetWs
            lo.ListRows(rowIdx).Delete
            RelockSheet targetWs
            deleted = deleted + 1
        Else
            missing = missing + 1
        End If
    Next i
    AppOn

    detail = deleted & " record" & IIf(deleted = 1, "", "s") & " deleted from the Reports page."
    If missing > 0 Then detail = detail & " " & missing & " could not be found (already removed?)."
    LogAudit "Delete visible (Reports)", "(multiple rooms)", detail

    Say detail, "A note of what was removed has been kept in the workbook's audit log."
    Exit Sub
Fail:
    AppReset
    ReportError "Delete visible records"
End Sub

Private Function FindReportRow(ByVal lo As ListObject, ByVal JobId As String) As Long
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Job ID").Value)), JobId, vbTextCompare) = 0 Then
            FindReportRow = i
            Exit Function
        End If
    Next i
End Function

Private Function ColByHeader(ByVal ws As Worksheet, ByVal HdrRow As Long, ByVal Header As String) As Long
    Dim c As Long
    For c = 1 To 100
        If Len(Trim$(CStr(ws.Cells(HdrRow, c).Value))) = 0 Then Exit Function
        If StrComp(CStr(ws.Cells(HdrRow, c).Value), Header, vbTextCompare) = 0 Then
            ColByHeader = c
            Exit Function
        End If
    Next c
End Function

' Snag list item 2d's minimum-columns view: hides every results column
' except this documented set, by header text against row 15 - the same
' "hide by name, touch nothing else" technique modInit.ApplyColumnVisibility
' uses for the location sheets' reduced-clutter view, reimplemented here
' rather than called directly because the results table is a spilled array
' with a header row, not an Excel Table/ListObject. The list itself is a
' fixed array, not a setting - unlike the location-sheet version, this one
' isn't user-facing yet (documented future enhancement, not built now).
Private Sub ApplyReportsMinimumColumns(ByVal ws As Worksheet, ByVal hdrs As Variant)
    Dim keep As Variant, i As Long, j As Long, c As Long, keepIt As Boolean
    keep = Array("Date/Time", "Location", "Student name", "Student no", "Paper stock", "Chargeable", "Paid")

    UnlockSheet ws
    For i = LBound(hdrs) To UBound(hdrs)
        c = ColByHeader(ws, 15, CStr(hdrs(i)))
        If c > 0 Then
            keepIt = False
            For j = LBound(keep) To UBound(keep)
                If StrComp(CStr(hdrs(i)), CStr(keep(j)), vbTextCompare) = 0 Then
                    keepIt = True
                    Exit For
                End If
            Next j
            ws.Columns(c).Hidden = Not keepIt
        End If
    Next i
    RelockSheet ws
End Sub

' ============================================================== helpers ===
Private Function SheetNamed(ByVal Nm As String, ByVal Position As Long) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Nm)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = Nm
    End If

    ' Move Before: is only correct when the sheet currently sits LATER than
    ' the target. Moving a sheet that is already earlier "before" the sheet at
    ' Position is a no-op - taking it out shifts everything after it down by
    ' one, so it lands back where it started. That is why every build left
    ' Reports on tab 1 and Summary on tab 2, the reverse of what this function
    ' asks for: the .xlsx ships neither report sheet, so build.ps1 always
    ' takes the first-run path, where Reports is created at index 1 and then
    ' asked to move to 2. Moving later needs After:.
    If ws.Index > Position Then
        ws.Move Before:=ThisWorkbook.Worksheets(Position)
    ElseIf ws.Index < Position Then
        ws.Move After:=ThisWorkbook.Worksheets(Position)
    End If

    Set SheetNamed = ws
End Function

' A spill cannot be written into a range that still holds an old one, and
' Cells.Clear does not always release a stale spill parent on a rebuild.
Private Sub ClearSpillArea(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range("A10:BZ5000").ClearContents
    On Error GoTo 0
End Sub

Private Sub TotalCell(ByVal ws As Worksheet, ByVal ValueAddr As String, ByVal Label As String, ByVal f As String)
    Dim lab As Range
    Set lab = ws.Range(ValueAddr).Offset(0, -1)
    lab.Value = Label
    lab.Font.Bold = True
    lab.HorizontalAlignment = xlRight
    ws.Range(ValueAddr).Formula2 = f
    ws.Range(ValueAddr).Font.Bold = True
End Sub

' Reports' own "Matching" totals (BuildReports): label directly above its
' value, same column, rather than TotalCell's label-to-the-left - see that
' call site's 2026-09-22 comment for why (a shared-column collision with
' the results table's hideable columns, §8.3).
Private Sub MatchTotal(ByVal ws As Worksheet, ByVal ColLetter As String, ByVal Label As String, ByVal f As String)
    ws.Range(ColLetter & "12").Value = Label
    ws.Range(ColLetter & "12").Font.Bold = True
    ws.Range(ColLetter & "13").Formula2 = f
    ws.Range(ColLetter & "13").Font.Bold = True
End Sub

' Label + input formatting for one filter box. Used to also write a static
' hint sentence one column right of the input ("Choose from the list, or
' leave blank for all.", etc.) - removed 2026-09-26: a sentence a user reads
' once and never again is pure visual clutter permanently, not help, and it
' was also the whole reason these rows needed wrapping/AutoFit in the first
' place. The same guidance still exists as an on-demand tooltip (AddList's
' own Title/Msg, AddDateValidation) that shows only when the cell is
' actually selected, instead of taking up screen space forever.
Private Sub CritCell(ByVal ws As Worksheet, ByVal LabelAddr As String, ByVal InputAddr As String, _
                     ByVal Label As String)
    ws.Range(LabelAddr).Value = Label
    ws.Range(LabelAddr).Font.Bold = True
    ws.Range(LabelAddr).WrapText = True
    With ws.Range(InputAddr)
        .Locked = False
        .Interior.Color = RGB(255, 255, 255)
        .Borders(xlEdgeLeft).Color = RGB(46, 100, 168)
        .Borders(xlEdgeLeft).Weight = xlMedium
    End With
    ' No AutoFit here (deliberately - tried, then reverted). Columns A/D/N
    ' are still at Excel's factory-default width at this point in the build:
    ' the Reports sheet is recreated from scratch every run (SheetNamed adds
    ' it fresh when missing), and FormatReports's own ColumnWidth calls, much
    ' later in BuildReports, are the first thing to widen them. BuildReports
    ' does the one real AutoFit pass itself, right after FormatReports sets
    ' the widths this depends on.
End Sub

Private Sub WriteHeaderRow(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal Headers As Variant)
    Dim i As Long
    For i = LBound(Headers) To UBound(Headers)
        ws.Cells(RowNo, i - LBound(Headers) + 1).Value = Headers(i)
    Next i
End Sub

' ==================================================== export report sig ===
' A signature of the Reports sheet's CURRENT filtered/sorted results - same
' shape as modExport.ExportSig (row count, cell count, a cost sum, a max
' date) - so "Export report" can stamp what it captured, and the Reports-page
' delete can warn when the visible set has moved on since.
Public Function FilteredSig(ByVal ws As Worksheet) As String
    Dim rng As Range, n As Long, cellCount As Double, chg As Double, last As Double, bad As String

    On Error Resume Next
    Set rng = ws.Range("A16").SpillingToRange
    On Error GoTo 0
    If rng Is Nothing Then
        FilteredSig = "empty"
        Exit Function
    End If

    ' A16 may be spilling the "no jobs match"/"no jobs recorded" fallback
    ' TEXT rather than real rows. Row count alone reads 1 either way, but the
    ' first cell of a real row is always a Date/Time serial number, so
    ' ISNUMBER is what actually tells the two apart.
    If Not IsNumeric(rng.Cells(1, 1).Value2) Then
        FilteredSig = "empty"
        Exit Function
    End If

    n = rng.Rows.Count
    cellCount = Application.WorksheetFunction.CountA(rng)
    chg = ResultsAgg(rng, 12, "Sum", bad)   ' column L = Chargeable
    last = ResultsAgg(rng, 1, "Max", bad)   ' column A = Date/Time

    FilteredSig = n & "|" & Format$(cellCount, "0") & "|" & Format$(chg, "0.####") & _
                  "|" & Format$(last, "0.######") & bad
End Function

Private Function ResultsAgg(ByVal rng As Range, ByVal ColOffset As Long, ByVal What As String, ByRef Bad As String) As Double
    On Error GoTo Failed
    Select Case What
        Case "Sum": ResultsAgg = Application.WorksheetFunction.Sum(rng.Columns(ColOffset))
        Case "Max": ResultsAgg = Application.WorksheetFunction.Max(rng.Columns(ColOffset))
        Case Else: GoTo Failed
    End Select
    Exit Function
Failed:
    ResultsAgg = 0
    Bad = Bad & "|?" & What
End Function

Public Function ReportsSheet() As Worksheet
    On Error Resume Next
    Set ReportsSheet = ThisWorkbook.Worksheets(SHEET_REPORTS)
End Function

Public Function ReportsExportSig(ByVal ws As Worksheet) As String
    ReportsExportSig = CStr(ws.Range(EXPORT_SIG_CELL).Value)
End Function

' Called by modExport.ExportReportSnapshot once the .xlsx has actually been
' written - never optimistically, for the same reason modExport.StampExported
' only runs after modExport confirms the file exists.
Public Sub StampReportsExport(ByVal ws As Worksheet)
    UnlockSheet ws
    ws.Range(EXPORT_SIG_CELL).Value = FilteredSig(ws)
    ws.Range(EXPORT_WHEN_CELL).Value = Now
    ws.Range(EXPORT_WHEN_CELL).NumberFormat = "dd/mm/yyyy hh:mm"
    RelockSheet ws
End Sub

' A dropdown from a literal comma-separated Formula1 list, e.g. Ascending,
' Descending - fine here because both lists (report headers, Asc/Desc) are
' short and fixed, well under xlValidateList's 255-character literal limit.
Private Sub AddList(ByVal target As Range, ByVal QuotedItems As String, ByVal Title As String, ByVal Msg As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Replace$(QuotedItems, """", "")
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = "Choose one of the listed options, or leave it blank."
    End With
End Sub

' Date-type validation, range-checked on manual entry, with no ActiveX
' involved - consistent with modPicker's Mac-safe, no-ActiveX rule. Validation
' only fires on manual entry, never on a value set from VBA/COM, so this adds
' a UI convenience without narrowing what modImport or a test script can
' write, and the Criteria() formula's own tolerance for a text date left by an
' unformatted cell is still needed and unchanged.
'
' This was written expecting Date validation to also draw Excel's small
' calendar icon on selection, the way it does in Excel for the web. Checked
' against the 2026-09-21 snag list item 5 report (no icon on the latest
' Excel, Windows or Mac): that icon has only ever shipped to Excel for the
' web, with no announced desktop date, so its absence here is a platform gap
' rather than a bug. A custom worksheet-based picker (the same no-ActiveX
' pattern modPicker.bas uses for the printer/paper multi-select) would give
' desktop Excel a working calendar, but was deferred rather than built.
'
' The lower bound is a real DATE() formula, not a literal date string - a
' string like "01/01/2000" is read back through the machine's locale, the
' same class of trap modUtils.NumOf's comment warns about for numbers.
Private Sub AddDateValidation(ByVal target As Range, ByVal Title As String, ByVal Msg As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="=DATE(2000,1,1)"
        .IgnoreBlank = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = "Enter a valid date on or after 1 January 2000, or leave blank."
    End With
End Sub

' Technician dropdown, sourced from what has actually been recorded (the
' consolidated _Data range) rather than the current Technicians catalogue -
' so a job against a technician no longer active is still findable, and
' every choice offered is guaranteed to match at least one record.
'
' Printer and Paper Stock are different (snag list item 4): they list every
' active printer/stock DEFINED in the catalogue, not just what has been used
' yet, so a printer or stock added today is immediately choosable here. The
' two lists are independent - deliberately not cross-filtered by printer
' compatibility the way BindStockRange does for job entry - so picking an
' incompatible pair just matches no rows, same as any other filter combination
' that matches nothing.
'
' Public, and also called from modRegistry.RefreshLocations after it
' rewrites _Data - not just from here. _Data does not exist yet the first
' time BuildReports runs (InitialiseWorkbook builds the report sheets before
' RefreshLocations ever writes it), and it changes on every later refresh
' too, so a one-time snapshot taken only at build time would read empty on
' a fresh workbook and go stale the moment a job is added anywhere.
Public Sub RefreshReportFilterLists(ByVal ws As Worksheet)
    ' Student name/no (snag list item 3): a dropdown of what has been
    ' recorded, for browsing and autocomplete - Strict:=False so it never
    ' blocks a fragment search or a number not yet logged.
    ApplyTo ws, ws.Range("B4"), DistinctValues("Student Name"), "REP|StudentName", _
        "Student/Department name", "Pick a recorded name, or type any text - part of a name is enough.", Strict:=False
    ApplyTo ws, ws.Range("B5"), DistinctValues("Student No"), "REP|StudentNo", _
        "Student/Department number", "Pick a recorded number, or type one that hasn't been logged yet.", Strict:=False

    ' Rows shifted down one, F5:F8 (2026-09-27) - see BuildReports' CritCell
    ' call sites for why.
    ApplyTo ws, ws.Range("F5"), DistinctValues("Technician"), "REP|Technician", _
        "Technician", "Choose a technician, or leave blank to include all."
    ApplyTo ws, ws.Range("F6"), AllActivePrinters(), "REP|Printer", _
        "Printer", "Choose a printer, or leave blank to include all."
    ApplyTo ws, ws.Range("F7"), AllActiveStocks(), "REP|Paper stock", _
        "Paper stock", "Choose a paper stock, or leave blank to include all."

    ' Location (print room): every registered print room (modRegistry.
    ' AllLocationCodes, tblLocations), not just those with a job logged yet -
    ' same "every catalogue entry" behaviour as Printer/Paper Stock above
    ' (AllActivePrinters/AllActiveStocks), so a room added today is
    ' immediately choosable here. Moved from O4 to F4 (2026-09-27) - see
    ' BuildReports' CritCell call site.
    ApplyTo ws, ws.Range("F4"), AllLocationCodes(), "REP|Location", _
        "Location (print room)", "Choose a print room, or leave blank to include all."
End Sub

' Distinct, sorted, non-blank values of one column of the consolidated range,
' read directly in VBA rather than as a formula - ApplyTo's staging column
' needs a Collection of plain values, not a spilled array.
Private Function DistinctValues(ByVal Header As String) As Collection
    Dim out As Collection, seen As clsDict
    Dim rng As Range, hdrRng As Range, m As Variant
    Dim hdrCol As Long, i As Long, j As Long, v As String, tmp As String
    Dim vals() As String, n As Long

    Set out = New Collection
    Set rng = ConsolidatedRange()
    If rng Is Nothing Then
        Set DistinctValues = out
        Exit Function
    End If

    ' DATA_HDR is already sheet-qualified ("_Data!$A$9:$AZ$9"), and
    ' Worksheet.Range does not reliably accept a qualified address string -
    ' Application.Range does.
    Set hdrRng = Application.Range(DATA_HDR)
    m = Application.Match(Header, hdrRng, 0)
    If IsError(m) Then
        Set DistinctValues = out
        Exit Function
    End If
    hdrCol = CLng(m)   ' Both rng and hdrRng start at column A, so this offset lines up with rng directly.

    Set seen = New clsDict
    ReDim vals(1 To rng.Rows.Count)
    For i = 1 To rng.Rows.Count
        v = Trim$(CStr(rng.Cells(i, hdrCol).Value))
        If Len(v) > 0 Then
            If Not seen.Exists(v) Then
                seen.Add v, True
                n = n + 1
                vals(n) = v
            End If
        End If
    Next i

    ' An insertion sort is plenty for a list this size (printers, paper
    ' stocks, technicians number in the tens, not thousands).
    For i = 2 To n
        tmp = vals(i)
        j = i - 1
        Do While j >= 1
            If StrComp(vals(j), tmp, vbTextCompare) <= 0 Then Exit Do
            vals(j + 1) = vals(j)
            j = j - 1
        Loop
        vals(j + 1) = tmp
    Next i

    For i = 1 To n
        out.Add vals(i)
    Next i
    Set DistinctValues = out
End Function
