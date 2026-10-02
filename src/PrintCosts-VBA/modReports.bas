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

' The results table's data rows, for the editable Paid column (see "edit Paid
' on Reports" below). 2000 matches the number formats FormatReports applies.
Private Const RESULTS_FIRST_ROW As Long = 18
Private Const RESULTS_LAST_ROW As Long = 2000

' Which job a results-table Paid cell showed when it was selected - see "edit
' Paid on Reports". Declared here, with the module's other declarations, not
' beside the code that uses it (a module-level declaration after the first
' Sub has failed to compile in this project before - see modInit's comment).
Private Type PaidSnap
    Addr As String
    Job As String
    Loc As String
    Paid As String
End Type
Private mCur As PaidSnap
Private mPrev As PaidSnap

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
    s = s & "*IF($F$7="""",TRUE," & C("Printer") & "=$F$7)"
    s = s & "*IF($F$8="""",TRUE," & C("Paper Stock") & "=$F$8)"

    ' Paper type (Roll/Sheet, 2026-10-02): read off the consolidated Unit
    ' column - "sheets" for a sheet job, "metres" for a roll job (a
    ' centimetres room's rows are already converted to metres in _Data,
    ' modRegistry.MetresBlock). A job with no Unit yet (no stock chosen)
    ' matches neither choice. (A Quantity filter that used to sit below
    ' Paper stock was removed the same day, direct user request.)
    s = s & "*IF($F$9="""",TRUE,IF($F$9=""Sheet""," & C("Unit") & "=""sheets"",ISNUMBER(MATCH(" & C("Unit") & ",{""metres"",""cm""},0))))"

    ' Paid (Yes/No, 2026-10-02): No means "not marked Yes", so a blank Paid
    ' (a job that predates the column, or never marked) counts as unpaid,
    ' the same reconciliation Matching's Paid/Unpaid totals use. INDEX over a
    ' genuinely blank cell reads back as 0, so compare against "Yes" rather
    ' than against "No".
    s = s & "*IF($B$10="""",TRUE,IF($B$10=""Yes""," & C("Paid") & "=""Yes""," & C("Paid") & "<>""Yes""))"

    ' Location (print room), added 2026-09-25 - a dropdown of registered print
    ' rooms (RefreshReportFilterLists, AllLocationCodes), same exact-match
    ' treatment as Technician/Printer/Paper Stock above. Moved from $O$4 to
    ' $F$4 (2026-09-27, see its CritCell call site) once Technician/Printer/
    ' Paper Stock shifted down a row and freed it.
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

    ' "Mark all as..." cluster (2026-09-29, direct user request): 1 column x 3
    ' rows at O1:O3, in the header strip just left of Export report / Delete
    ' visible records (T). The label is cell text here; the Paid and Unpaid
    ' buttons for rows 2 and 3 are drawn by modInit.InitialiseWorkbook, which
    ' runs after this so the row heights they are sized to are already
    ' settled.
    '
    ' Why O, found the hard way (first tried V, beside the T buttons): (1)
    ' test-reports.ps1 requires every Reports button to sit within the first
    ' screenful (Left <= 900pt) and V is at ~1020pt. (2) A2's instruction text
    ' runs to ~689pt, so anything at N or earlier would sit on top of its
    ' tail; O starts at ~720pt, clear of it. (3) O is "Paid" in the results
    ' table, one of the columns ApplyReportsMinimumColumns always keeps
    ' visible, so the cluster cannot vanish with a hidden column - the same
    ' reason the (since removed) Export names toggle lived at N:O.
    ws.Range("O1").Value = "Mark all as..."
    ws.Range("O1").Font.Bold = True

    ' Label | input | hint | gap, in that order (snag list item 2) - the
    ' input sits immediately right of its label, and the unused column at the
    ' end of each group is what separates it from the next.
    ' "Student/Department" labels (snag 4b, 2026-09-22): Student Name/No is
    ' also used free-text for department charging (D19) - the underlying
    ' column headers stay "Student Name"/"Student No" (renaming those would
    ' break every header-name lookup that already reads them, for no real
    ' capability gained), but the on-screen wording someone actually reads
    ' here reads sensibly either way.
    '
    ' Filter block layout (2026-10-02, direct user request) - two columns,
    ' with a blank row between each group:
    '   rows 4-5   A:B  Student/dept. name, Student number
    '                   D:F  Location, Technician
    '   row 6      gap
    '   rows 7-8   A:B  From date, To date
    '   rows 7-9   D:F  Printer, Paper stock, Paper type
    '   row 9      gap (left)
    '   row 10     A:B  Paid
    '   row 11     gap, then Sort by at row 12
    CritCell ws, "A4", "B4", "Student/dept. name"
    CritCell ws, "A5", "B5", "Student number"
    CritCell ws, "A7", "B7", "From date"
    CritCell ws, "A8", "B8", "To date"
    ws.Range("B7:B8").NumberFormat = "dd/mm/yyyy"
    AddDateValidation ws.Range("B7"), "From date", "Leave blank for no start date."
    AddDateValidation ws.Range("B8"), "To date", "Jobs logged at any time on this date are included."

    ' Paid (Yes/No): Yes = only jobs marked Paid; No = everything not marked
    ' Paid, blank included. Left empty to ignore it.
    CritCell ws, "A10", "B10", "Paid"
    AddList ws.Range("B10"), """Yes"",""No""", "Paid", "Yes shows only paid jobs, No only unpaid ones (including jobs never marked). Leave blank to include both."

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
    ' Technician/Printer/Paper stock moved down one row, D5:D7/F5:F7
    ' (2026-09-27, direct user request), freeing D4/F4 for the Location
    ' (print room) filter immediately below - see that CritCell call for why
    ' it moved out of N4/O4.
    CritCell ws, "D5", "F5", "Technician"
    ' Row 6 is a gap (2026-10-02), then Printer / Paper stock / Paper type.
    CritCell ws, "D7", "F7", "Printer"
    CritCell ws, "D8", "F8", "Paper stock"
    CritCell ws, "D9", "F9", "Paper type"
    AddList ws.Range("F9"), """Roll"",""Sheet""", "Paper type", "Roll shows only roll jobs, Sheet only sheet jobs. Leave blank to include both."

    ' Location (print room) filter. Originally parked at N4/O4 (2026-09-25) to
    ' dodge E/G/H/I:M, the columns ApplyReportsMinimumColumns hides entirely
    ' by results-header name (the same trap O10's own 2026-09-22 comment
    ' names) - a label or input parked there can render hidden or orphaned
    ' under the default view. Moved to D4/F4 (2026-09-27, direct user
    ' request) - the row the Technician/Printer/Paper Stock group
    ' vacated by shifting down one row (above) - rather than staying at N4/O4,
    ' so the whole filter block reads top-to-bottom as one contiguous group:
    ' Location, Technician, Printer, Paper stock.
    CritCell ws, "D4", "F4", "Location (print room)"

    ' Sort by/direction sit below the filters, above the totals row (snag list
    ' item 3) rather than beside the filters - row 11 is the "name and number
    ' don't match" warning below, the gap between the filters and the sort
    ' settings, and row 13 is the gap before Matching. Moved from row 10 to
    ' row 12 (2026-10-02) when the filter block grew to row 10.
    CritCell ws, "A12", "B12", "Sort by"
    CritCell ws, "D12", "F12", "Sort direction"
    AddList ws.Range("B12"), QuotedList(hdrs), "Sort by", "Which column to sort the results by."
    AddList ws.Range("F12"), """Ascending"",""Descending""", "Sort direction", "Which way to sort."

    ' Export names (a Yes/No toggle that lived at N12/O12) was removed
    ' 2026-10-02, direct user request: Export report now asks Yes / No /
    ' Cancel at export time instead (modExport.AskIncludeNamesReport), the
    ' same as the room exports. The live results always show names.

    ' Spec 14.1: both criteria given, neither matching the other. Row 11
    ' (was row 9 before 2026-10-02): the gap row under the filters.
    ws.Range("A11").Formula2 = "=IF(OR($B$4="""",$B$5=""""),""""," & _
        "IF(IFERROR(ROWS(FILTER(" & C("Job ID") & "," & ok & ")),0)=0," & _
        """That name and that number do not appear together on any record - check both."",""""))"
    ws.Range("A11").Font.Color = RGB(176, 0, 32)
    ws.Range("A11").Font.Bold = True

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
    ws.Range("A14").Value = "Matching"
    ws.Range("A14").Font.Bold = True
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
    ws.Range("C15").NumberFormat = CurrencyFormatCode()
    ws.Range("D15").NumberFormat = CurrencyFormatCode()
    ws.Range("F15").NumberFormat = CurrencyFormatCode()
    ws.Range("N15").NumberFormat = CurrencyFormatCode()
    ws.Range("O15").NumberFormat = CurrencyFormatCode()

    ' --- the records ------------------------------------------------------
    ' Job ID is appended after Notes and hidden - the correlation key that
    ' maps a visible row back to its source location sheet and table row for
    ' the Reports-page delete. Nothing else moves, so existing column
    ' positions are untouched.
    WriteHeaderRow ws, 17, hdrs
    ws.Cells(17, UBound(hdrs) - LBound(hdrs) + 2).Value = "Job ID"

    ' Sorting: the whole FILTER result is bound to res once via LET, then
    ' re-ordered by whichever column K5 names, found by matching its header
    ' text against the same header list the table itself uses so the two
    ' can never drift apart. IFERROR around the SORTBY step is what keeps
    ' FILTER's own "no rows matched"/"no data at all" fallback text intact
    ' when there is nothing to sort - SORTBY on that text would otherwise
    ' error and the outer IFERROR would show the wrong one of the two
    ' messages.
    ' Student Name/No: always shown here. The old "Export names" toggle
    ' (snag 2a) used to blank these values in this formula, which hid them
    ' from the live view too; since 2026-09-29 it is applied at export time
    ' instead (modExport.RedactNameColumn; since 2026-10-02 the export asks Yes/No instead), so this formula has no reference
    ' to a toggle at all.
    '
    ' Built with Chr(34) for the formula's own empty-string literal rather
    ' than hand-counting doubled quotes in a VBA string literal - a single
    ' wrong quote count here is a silent formula-text bug, not a compile
    ' error, so it is worth avoiding the manual counting entirely.
    '
    ' History, kept as a warning: when the O10 toggle lived in this formula,
    ' IF($O$10="Yes", <column array>, "") was NOT an elementwise blank - the
    ' CONDITION was a bare scalar, so with O10="No" the whole IF collapsed to
    ' the single scalar "" rather than a column of blanks the same height as
    ' every other HSTACK argument, and HSTACK does not broadcast a scalar
    ' against a column (modRegistry's own §7.1 comment names the same trap).
    ' The per-row IF(Job ID <> "", ...) wrapper below is what made that work
    ' and is kept as-is; do not reintroduce a scalar-conditioned IF here.
    Dim q As String, rowShape As String, sName As String, sNo As String
    q = Chr(34)
    rowShape = C("Job ID") & "<>" & q & q
    sName = "IF(" & rowShape & "," & C("Student Name") & "," & q & q & ")"
    sNo = "IF(" & rowShape & "," & C("Student No") & "," & q & q & ")"

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
    f = f & "sortIdx,IFERROR(MATCH($B$12,hdrs,0),0),"
    f = f & "dir,IF($F$12=""Descending"",-1,1),"
    f = f & "IF(sortIdx=0,res,IFERROR(SORTBY(res,INDEX(res,,sortIdx),dir),res))"
    f = f & "),""No print jobs have been recorded yet."")"
    ws.Range("A18").Formula2 = f
    ws.Columns(UBound(hdrs) - LBound(hdrs) + 2).Hidden = True

    ws.Range(EXPORT_SIG_CELL).Value = savedSig
    If IsDate(savedWhen) Then
        ws.Range(EXPORT_WHEN_CELL).Value = savedWhen
        ws.Range(EXPORT_WHEN_CELL).NumberFormat = "dd/mm/yyyy hh:mm"
    End If
    ws.Range(EXPORT_SIG_CELL).EntireColumn.Hidden = True

    BuildBreakdowns ws, ok
    FormatReports ws
    MakePaidEditable ws

    ' The filter labels/hints (CritCell, above) were written and wrapped
    ' before this point, while columns A/C/D/N still sat at Excel's
    ' factory-default width - FormatReports's ColumnWidth calls, just above,
    ' are the first thing to widen them. Re-fitting now, with the real widths
    ' in place, is what keeps rows 4:10 sized for the text that actually
    ' fits per line rather than for the cramped default - see CritCell's own
    ' comment for how this was found.
    ws.Rows("4:12").AutoFit

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
    ws.Rows("4:12").Group
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
    ws.Range("T17").Value = "By print room"
    ws.Range("T17").Font.Bold = True
    ws.Range("T18").Formula2 = GroupFormula(ok, "Location")

    ws.Range("X17").Value = "By paper stock"
    ws.Range("X17").Font.Bold = True
    ws.Range("X18").Formula2 = GroupFormula(ok, "Paper Stock")

    ' Each block spills as key | Jobs | Gross | Chargeable, so the money
    ' columns are the third and fourth - Jobs is a count and must not be
    ' formatted as currency.
    ws.Range("U19:U2000").NumberFormat = "#,##0"
    ws.Range("V19:W2000").NumberFormat = CurrencyFormatCode()
    ws.Range("Y19:Y2000").NumberFormat = "#,##0"
    ws.Range("Z19:AA2000").NumberFormat = CurrencyFormatCode()
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
    ws.Range("A14:O14").Interior.Color = RGB(244, 232, 222)
    ws.Range("A17:Q17").Interior.Color = RGB(222, 232, 244)
    ws.Range("A18:A2000").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range("G18:I2000").NumberFormat = "#,##0.00"
    ws.Range("J18:N2000").NumberFormat = CurrencyFormatCode()
    ws.Columns("A:Q").ColumnWidth = 14
    ws.Columns("B:F").ColumnWidth = 22
    ws.Columns("Q").ColumnWidth = 30
    ws.Columns("A").ColumnWidth = ColWidthForPx(180)  ' Date/Time, 180px
    ws.Columns("T").ColumnWidth = 22
    ws.Columns("X").ColumnWidth = 22
    ws.Rows(17).Font.Bold = True
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

    ' Filter safeguard (2026-09-29, direct user request) - see
    ' RequireActiveFilter. Checked before anything is read or counted, so an
    ' unfiltered sheet is refused outright rather than after a prompt built
    ' from every job in the workbook.
    If Not RequireActiveFilter(ws, "delete records") Then Exit Sub

    On Error Resume Next
    Set rng = ws.Range("A18").SpillingToRange
    On Error GoTo 0
    If rng Is Nothing Then
        Say "There is nothing to delete.", "The Reports sheet has no results under the current filters."
        Exit Sub
    End If
    ' A18 may be spilling FILTER's own "no jobs match"/"no jobs recorded"
    ' fallback TEXT rather than real rows - see FilteredSig's own comment.
    If Not IsNumeric(rng.Cells(1, 1).Value2) Then
        Say "There is nothing to delete.", CStr(rng.Cells(1, 1).Value)
        Exit Sub
    End If
    n = rng.Rows.Count

    locCol = ColByHeader(ws, 17, "Location")
    jobCol = ColByHeader(ws, 17, "Job ID")
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

    ' Same safeguard as the interactive entry point, repeated here because
    ' this is the routine that actually deletes (and the one a test or a
    ' future caller reaches directly, bypassing the Ask() gate).
    If Not RequireActiveFilter(ws, "delete records") Then Exit Sub

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
    ResnapReportsSelection   ' see "edit Paid here"

    detail = deleted & " record" & IIf(deleted = 1, "", "s") & " deleted from the Reports page."
    If missing > 0 Then detail = detail & " " & missing & " could not be found (already removed?)."
    LogAudit "Delete visible (Reports)", "(multiple rooms)", detail

    Say detail, "A note of what was removed has been kept in the workbook's audit log."
    Exit Sub
Fail:
    AppReset
    ReportError "Delete visible records"
End Sub

' ======================================================= edit Paid here ===
' (2026-09-29, direct user request) The Paid cells of the results table take
' the same Yes/No dropdown as the room sheets, and a choice is written back to
' the job record the row belongs to.
'
' The catch is that the results are ONE spilled formula. Typing into a cell of
' a spill range puts a constant there, turns the anchor (A18) into #SPILL!,
' and so blanks every cell of the table - including the row's Job ID and
' Location, which are exactly what identify the record. By the time the Change
' event fires they can no longer be read from the sheet. So:
'
'   - RememberReportsPaidCell runs whenever a cell is selected and notes which
'     job the selected Paid cell shows (and keeps the PREVIOUS selection's
'     note too, because Excel can move the selection after Enter either
'     before or after it raises Change - both orders are handled).
'   - OnReportsPaidEdited, from Workbook_SheetChange, applies the edit to that
'     job, then clears the typed constant so the spill comes back.
'
' Anything it cannot attribute to a real record - a cell below the results, a
' pasted block, a value that is not Yes/No, a record that changed since it was
' drawn - is cleared and refused, never left behind: an orphan constant in the
' results area would block the spill the next time the results grew that far.
'
' Two costs, both inherent: the sheet's Undo history is cleared by the write
' (Ctrl+Z will not undo a Paid change - change it back instead), and only one
' cell can be edited at a time (Mark all as... is the bulk route).
'
' The Paid cells are unlocked and given the dropdown by MakePaidEditable; the
' rest of the results table stays locked.
Private Sub MakePaidEditable(ByVal ws As Worksheet)
    Dim c As Long, rng As Range
    c = ColByHeader(ws, 17, "Paid")
    If c = 0 Then Exit Sub
    Set rng = ws.Range(ws.Cells(RESULTS_FIRST_ROW, c), ws.Cells(RESULTS_LAST_ROW, c))
    rng.Locked = False
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Paid"
        .InputMessage = "Choose Yes or No. This updates the job record itself, in its print room, straight away. Blank means not recorded and counts as unpaid."
        .ErrorTitle = "Paid"
        .ErrorMessage = "Choose Yes or No."
    End With
End Sub

' Notes which job a selected Paid cell shows. Called on every selection change
' on Reports, so the early exits come first and cost almost nothing.
Public Sub RememberReportsPaidCell(ByVal ws As Worksheet, ByVal Target As Range)
    mPrev = mCur
    TakeSnapshot ws, Target
End Sub

' Re-reads the note for the active cell without demoting the current one - for
' after something changed the results underneath a selection that has not
' moved (an edit, Mark all as..., Delete visible). The previous note is
' dropped: it described the results as they were.
Public Sub ResnapReportsSelection()
    Dim ws As Worksheet
    Dim blank As PaidSnap
    On Error Resume Next
    Set ws = ReportsSheet()
    If ws Is Nothing Then Exit Sub
    mPrev = blank
    If ActiveSheet Is ws Then TakeSnapshot ws, ActiveCell
    On Error GoTo 0
End Sub

Private Sub TakeSnapshot(ByVal ws As Worksheet, ByVal Target As Range)
    Dim blank As PaidSnap, a As Variant, locCol As Long, jobCol As Long
    mCur = blank
    If Target Is Nothing Then Exit Sub
    If Target.Cells.Count <> 1 Then Exit Sub
    If Target.Row < RESULTS_FIRST_ROW Or Target.Row > RESULTS_LAST_ROW Then Exit Sub
    If StrComp(CStr(ws.Cells(17, Target.Column).Value), "Paid", vbTextCompare) <> 0 Then Exit Sub

    ' A real record on this row: Date/Time, column A, is a number. (Empty when
    ' the row is past the results, text for FILTER's "no jobs match" message,
    ' an error while a typed constant is blocking the spill.)
    a = ws.Cells(Target.Row, 1).Value2
    If IsError(a) Then Exit Sub
    If IsEmpty(a) Then Exit Sub
    If Not IsNumeric(a) Then Exit Sub

    locCol = ColByHeader(ws, 17, "Location")
    jobCol = ColByHeader(ws, 17, "Job ID")
    If locCol = 0 Or jobCol = 0 Then Exit Sub
    If IsError(ws.Cells(Target.Row, jobCol).Value2) Then Exit Sub

    mCur.Job = Trim$(CStr(ws.Cells(Target.Row, jobCol).Value2))
    If Len(mCur.Job) = 0 Then Exit Sub
    mCur.Loc = CStr(ws.Cells(Target.Row, locCol).Value2)
    If IsError(Target.Value2) Then
        mCur.Paid = ""
    Else
        mCur.Paid = Trim$(CStr(Target.Value2))
    End If
    mCur.Addr = Target.Address
End Sub

' Workbook_SheetChange calls this for EVERY edit on Reports, so it returns at
' once unless the edit touches the Paid column of the results table.
Public Sub OnReportsPaidEdited(ByVal ws As Worksheet, ByVal Target As Range)
    Dim c As Long, area As Range, hit As Range, v As Variant
    Dim snap As PaidSnap, found As Boolean
    Dim newV As String, oldV As String
    Dim targetWs As Worksheet, lo As ListObject, rowIdx As Long, cel As Range

    c = ColByHeader(ws, 17, "Paid")
    If c = 0 Then Exit Sub
    Set area = ws.Range(ws.Cells(RESULTS_FIRST_ROW, c), ws.Cells(RESULTS_LAST_ROW, c))
    Set hit = Intersect(Target, area)
    If hit Is Nothing Then Exit Sub

    ' From here the edit is ours, and however it ends the typed constant has
    ' to go so the spill can come back.
    On Error GoTo Fail

    If Target.Cells.Count > 1 Then
        RestoreResults ws, hit
        Say "Change one Paid cell at a time.", "Pasting or filling several cells is not supported here.", "To mark many records at once, set a filter and use Mark all as..."
        Exit Sub
    End If

    If StrComp(mCur.Addr, hit.Address, vbBinaryCompare) = 0 And Len(mCur.Job) > 0 Then
        snap = mCur
        found = True
    ElseIf StrComp(mPrev.Addr, hit.Address, vbBinaryCompare) = 0 And Len(mPrev.Job) > 0 Then
        snap = mPrev
        found = True
    End If
    If Not found Then
        RestoreResults ws, hit
        Say "That cell is not a job record, so nothing was changed.", "Only the Paid cell of a row that shows a record can be edited."
        Exit Sub
    End If

    v = hit.Value2
    If IsError(v) Then newV = "" Else newV = Trim$(CStr(v))
    If StrComp(newV, "Yes", vbTextCompare) = 0 Then
        newV = "Yes"
    ElseIf StrComp(newV, "No", vbTextCompare) = 0 Then
        newV = "No"
    Else
        RestoreResults ws, hit
        Say "Paid must be Yes or No.", "Nothing was changed.", "Pick Yes or No from the dropdown."
        Exit Sub
    End If

    ' Same as the row already showed: nothing to write, just put the spill back.
    oldV = snap.Paid
    If StrComp(newV, oldV, vbTextCompare) = 0 Then
        RestoreResults ws, hit
        Exit Sub
    End If

    Set targetWs = SheetForCode(snap.Loc)
    If Not targetWs Is Nothing Then Set lo = JobsTable(targetWs)
    If Not lo Is Nothing Then rowIdx = FindReportRow(lo, snap.Job)
    If rowIdx = 0 Then
        RestoreResults ws, hit
        Say "That record could not be found.", "It may have been removed or moved since the report was drawn. Nothing was changed."
        Exit Sub
    End If

    ' The guard against a stale note: the record's own Paid value must still be
    ' what this row showed when it was selected.
    Set cel = CellIn(lo, rowIdx, "Paid")
    If StrComp(Trim$(CStr(cel.Value)), oldV, vbTextCompare) <> 0 Then
        RestoreResults ws, hit
        Say "That record has changed since it was drawn.", "Its Paid value is no longer what this row showed, so nothing was changed.", "Look at the row again, then retry."
        Exit Sub
    End If

    AppOff
    UnlockSheet targetWs
    cel.Value = newV
    RelockSheet targetWs
    LogAudit "Paid edited (Reports)", snap.Loc, snap.Job & ": " & IIf(Len(oldV) = 0, "(blank)", oldV) & " -> " & newV
    RestoreResults ws, hit
    AppOn
    ResnapReportsSelection
    Exit Sub
Fail:
    ' Best effort to leave no constant behind, then report.
    On Error Resume Next
    If Not targetWs Is Nothing Then RelockSheet targetWs
    UnlockSheet ws
    hit.ClearContents
    RelockSheet ws
    On Error GoTo 0
    AppReset
    ReportError "Edit Paid"
End Sub

' Removes the constant a Paid edit left in the results area, so the spilled
' formula can fill it again, and re-notes the selection afterwards.
Private Sub RestoreResults(ByVal ws As Worksheet, ByVal rng As Range)
    AppOff
    UnlockSheet ws
    rng.ClearContents
    RelockSheet ws
    AppOn
    ResnapReportsSelection
End Sub

' ====================================================== filter safeguard ===
' "Delete visible records" and "Mark all as Paid/Unpaid" both act on every
' record the Reports sheet currently shows. With no filter set that is every
' job in the workbook, so both refuse to run until at least one filter is
' filled in (2026-09-29, direct user request). It is a guard against pressing
' the wrong button on the unfiltered sheet, not a substitute for the
' confirmation prompts, which still follow.
'
' Mirrors Criteria(), not merely "is the box non-empty": a box counts only if
' it actually narrows the results. The free-text and dropdown boxes count when
' they hold something other than spaces (Criteria itself would treat a lone
' space as a real search term - refusing is the safe direction). From date,
' To date are coerced with *1 in Criteria and IGNORED when that
' fails, so text that will not coerce does not count here either. Sort by and
' Sort direction are not filters and are not looked at.
'
' Cell addresses match Criteria: B4 name, B5 number, B7/B8 dates, F4 room,
' F5 technician, F7 printer, F8 paper stock, F9 paper type, B10 paid. If a filter box
' moves, change it here AND there.
Public Function HasActiveFilter(ByVal ws As Worksheet) As Boolean
    Dim a As Variant
    For Each a In Array("B4", "B5", "B10", "F4", "F5", "F7", "F8", "F9")
        If FilterBoxHasText(ws.Range(CStr(a))) Then
            HasActiveFilter = True
            Exit Function
        End If
    Next a
    For Each a In Array("B7", "B8")
        If FilterBoxHasText(ws.Range(CStr(a))) Then
            If Not CBool(ws.Evaluate("ISERROR(" & CStr(a) & "*1)")) Then
                HasActiveFilter = True
                Exit Function
            End If
        End If
    Next a
End Function

Private Function FilterBoxHasText(ByVal r As Range) As Boolean
    Dim v As Variant
    v = r.Value2
    If IsError(v) Then
        FilterBoxHasText = True
    Else
        FilterBoxHasText = (Len(Trim$(CStr(v))) > 0)
    End If
End Function

' True when a filter is set; otherwise says why the command will not run and
' returns False. Action completes "Set at least one filter before you ...".
Public Function RequireActiveFilter(ByVal ws As Worksheet, ByVal Action As String) As Boolean
    If HasActiveFilter(ws) Then
        RequireActiveFilter = True
        Exit Function
    End If
    Say "Set at least one filter before you " & Action & ".", _
        "With no filter set, every record in the workbook is shown, so this would change all of them. " & _
        "As a safeguard it only runs while at least one filter is filled in: student or department, dates, " & _
        "print room, technician, printer, paper stock, paper type or paid status.", _
        "Fill in a filter, check that the records shown are the ones you mean, then try again."
End Function

' ============================================================ mark paid ===
' "Mark all as..." Paid / Unpaid (2026-09-29, direct user request): sets the
' Paid column of every record the Reports sheet currently shows, and nothing
' else. Records the filters exclude are never touched, and neither are the
' filters themselves - this reads the results range and writes to the source
' room tables only, so the sheet's filter boxes and sort settings are exactly
' as the user left them.
'
' Same shape as the bulk delete above: an interactive entry point that
' confirms, and a Confirmed routine that does the work (Public so a test can
' call it - Ask() always declines under SetQuiet).
Public Sub MarkVisibleReports(ByVal NewPaid As String)
    Dim ws As Worksheet, rng As Range, n As Long
    Dim locCol As Long, jobCol As Long, paidCol As Long, i As Long
    Dim rooms As clsDict, k As Variant, breakdown As String
    Dim label As String, loc As String, same As Long, msg As String

    label = IIf(NewPaid = "Yes", "Paid", "Unpaid")

    Set ws = ReportsSheet()
    If ws Is Nothing Then
        Say "The Reports sheet could not be found.", "Run Refresh Locations first."
        Exit Sub
    End If

    If Not RequireActiveFilter(ws, "mark records as " & label) Then Exit Sub

    On Error Resume Next
    Set rng = ws.Range("A18").SpillingToRange
    On Error GoTo 0
    If rng Is Nothing Then
        Say "There is nothing to mark.", "The Reports sheet has no results under the current filters."
        Exit Sub
    End If
    ' A18 may be spilling FILTER's own "no jobs match"/"no jobs recorded"
    ' fallback TEXT rather than real rows - see FilteredSig's own comment.
    If Not IsNumeric(rng.Cells(1, 1).Value2) Then
        Say "There is nothing to mark.", CStr(rng.Cells(1, 1).Value)
        Exit Sub
    End If
    n = rng.Rows.Count

    locCol = ColByHeader(ws, 17, "Location")
    jobCol = ColByHeader(ws, 17, "Job ID")
    paidCol = ColByHeader(ws, 17, "Paid")
    If locCol = 0 Or jobCol = 0 Or paidCol = 0 Then
        Say "The Reports sheet layout looks wrong.", "The Location, Paid or Job ID column could not be found.", "Rebuild the report sheets (Refresh Locations), then try again."
        Exit Sub
    End If

    ' Per-room breakdown, as the delete prompt does, plus how many are
    ' already at the requested value (they are left alone, not rewritten).
    Set rooms = New clsDict
    For i = 1 To n
        loc = CStr(rng.Cells(i, locCol).Value)
        If rooms.Exists(loc) Then
            rooms.Add loc, CLng(rooms.Item(loc)) + 1
        Else
            rooms.Add loc, 1
        End If
        If StrComp(Trim$(CStr(rng.Cells(i, paidCol).Value)), NewPaid, vbTextCompare) = 0 Then same = same + 1
    Next i
    For Each k In rooms.Keys
        breakdown = breakdown & "  " & CStr(k) & ": " & rooms.Item(CStr(k)) & vbCrLf
    Next k

    msg = "Mark " & n & " visible record" & IIf(n = 1, "", "s") & " as " & label & "?" & vbCrLf & vbCrLf & breakdown
    If same > 0 Then msg = msg & vbCrLf & same & " already " & IIf(same = 1, "is", "are") & " marked " & label & " and will stay as they are."
    msg = msg & vbCrLf & vbCrLf & "Only the records shown are changed. Anything your filters hide is left alone, and the filters themselves stay as they are."

    If Not Ask(msg, "Mark visible records as " & label) Then Exit Sub

    MarkVisibleReportsConfirmed ws, rng, locCol, jobCol, n, NewPaid
End Sub

' The actual write. NewPaid is "Yes" or "No".
'
' Job IDs are indexed once per room (one read of the Job ID column) rather
' than located row by row with FindReportRow, which would rescan the whole
' room table for every visible record.
Public Sub MarkVisibleReportsConfirmed(ByVal ws As Worksheet, ByVal rng As Range, _
                                       ByVal LocCol As Long, ByVal JobCol As Long, ByVal n As Long, _
                                       ByVal NewPaid As String)
    Dim i As Long, r As Long, rowIdx As Long, cnt As Long
    Dim changed As Long, unchanged As Long, missing As Long
    Dim detail As String, label As String, loc As String
    Dim jobIds() As String, locs() As String, ids As Variant
    Dim targetWs As Worksheet, lo As ListObject, idCol As Range, cel As Range
    Dim maps As clsDict, los As clsDict, map As clsDict, k As Variant

    If NewPaid <> "Yes" And NewPaid <> "No" Then
        Err.Raise vbObjectError + 514, "MarkVisibleReportsConfirmed", "NewPaid must be ""Yes"" or ""No""."
    End If
    label = IIf(NewPaid = "Yes", "Paid", "Unpaid")

    ' Repeated here, as in DeleteVisibleReportsConfirmed - this is the routine
    ' that actually writes.
    If Not RequireActiveFilter(ws, "mark records as " & label) Then Exit Sub

    On Error GoTo Fail

    ' Copied into memory FIRST - rng is a live spilled formula range, and
    ' writing Paid recalculates it. See DeleteVisibleReportsConfirmed.
    ReDim jobIds(1 To n)
    ReDim locs(1 To n)
    For i = 1 To n
        jobIds(i) = CStr(rng.Cells(i, JobCol).Value)
        locs(i) = CStr(rng.Cells(i, LocCol).Value)
    Next i

    AppOff
    Set maps = New clsDict   ' room code -> (Job ID -> table row number)
    Set los = New clsDict    ' room code -> its jobs ListObject
    For i = 1 To n
        loc = locs(i)
        If Not maps.Exists(loc) Then
            Set map = New clsDict
            Set targetWs = SheetForCode(loc)
            If Not targetWs Is Nothing Then
                Set lo = JobsTable(targetWs)
                If Not lo Is Nothing Then
                    cnt = lo.ListRows.Count
                    If cnt > 0 Then
                        Set idCol = lo.ListColumns(ColIdx(lo, "Job ID")).DataBodyRange
                        ' A one-row range reads back as a scalar, not an array.
                        If cnt = 1 Then
                            ReDim ids(1 To 1, 1 To 1)
                            ids(1, 1) = idCol.Value2
                        Else
                            ids = idCol.Value2
                        End If
                        For r = 1 To cnt
                            map.Add Trim$(CStr(ids(r, 1))), r
                        Next r
                        los.Add loc, lo
                    End If
                End If
            End If
            maps.Add loc, map
        End If
    Next i

    ' Sheets unlocked once each, up front, and relocked together at the end
    ' (and in the error path) - not per record.
    For Each k In los.Keys
        UnlockSheet los.Obj(CStr(k)).Parent
    Next k

    For i = 1 To n
        loc = locs(i)
        rowIdx = 0
        Set map = maps.Obj(loc)
        If Not map Is Nothing Then
            If map.Exists(Trim$(jobIds(i))) Then rowIdx = CLng(map.Item(Trim$(jobIds(i))))
        End If
        If rowIdx = 0 Or Not los.Exists(loc) Then
            missing = missing + 1
        Else
            Set cel = CellIn(los.Obj(loc), rowIdx, "Paid")
            If StrComp(Trim$(CStr(cel.Value)), NewPaid, vbTextCompare) = 0 Then
                unchanged = unchanged + 1
            Else
                cel.Value = NewPaid
                changed = changed + 1
            End If
        End If
    Next i

    RelockRooms los
    AppOn
    ' The results have just changed under whatever cell is selected - see
    ' "edit Paid here".
    ResnapReportsSelection

    detail = changed & " record" & IIf(changed = 1, "", "s") & " marked " & label & " from the Reports page."
    If unchanged > 0 Then detail = detail & " " & unchanged & " already " & IIf(unchanged = 1, "was", "were") & " " & label & "."
    If missing > 0 Then detail = detail & " " & missing & " could not be found (already removed?)."
    LogAudit "Mark visible " & label & " (Reports)", "(multiple rooms)", detail

    Say detail, "Your filters have not been changed. A note has been kept in the workbook's audit log."
    Exit Sub
Fail:
    RelockRooms los
    AppReset
    ReportError "Mark visible records"
End Sub

' Relocks every room sheet MarkVisibleReportsConfirmed unlocked. Safe to call
' with nothing (the error path can be reached before los exists).
Private Sub RelockRooms(ByVal los As clsDict)
    Dim k As Variant
    If los Is Nothing Then Exit Sub
    On Error Resume Next
    For Each k In los.Keys
        RelockSheet los.Obj(CStr(k)).Parent
    Next k
    On Error GoTo 0
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
        c = ColByHeader(ws, 17, CStr(hdrs(i)))
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
    ws.Range(ColLetter & "14").Value = Label
    ws.Range(ColLetter & "14").Font.Bold = True
    ws.Range(ColLetter & "15").Formula2 = f
    ws.Range(ColLetter & "15").Font.Bold = True
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
    Set rng = ws.Range("A18").SpillingToRange
    On Error GoTo 0
    If rng Is Nothing Then
        FilteredSig = "empty"
        Exit Function
    End If

    ' A18 may be spilling the "no jobs match"/"no jobs recorded" fallback
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

    ' Rows re-laid out 2026-10-02 (F5, F7:F9) - see BuildReports' CritCell
    ' call sites for why.
    ApplyTo ws, ws.Range("F5"), DistinctValues("Technician"), "REP|Technician", _
        "Technician", "Choose a technician, or leave blank to include all."
    ApplyTo ws, ws.Range("F7"), AllActivePrinters(), "REP|Printer", _
        "Printer", "Choose a printer, or leave blank to include all."
    ApplyTo ws, ws.Range("F8"), AllActiveStocks(), "REP|Paper stock", _
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
