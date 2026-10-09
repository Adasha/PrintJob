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
' Summary table position. The job planner box (modPlanner, rows 5-21) sits above it,
' so these are 18 rows lower than they were before the planner existed (0.10.29).
Public Const SUM_TOT_ROW As Long = 24      ' totals values; their labels are the row above
Public Const SUM_HDR_ROW As Long = 27      ' table header
Public Const SUM_FIRST_ROW As Long = 28    ' first row of the spilled table
Private Const DATA_SPILL As String = "_Data!$A$10#"
Private Const DATA_HDR As String = "_Data!$A$9:$AZ$9"
Private Const EXPORT_SIG_CELL As String = "AN1"
Private Const EXPORT_WHEN_CELL As String = "AN2"

' The results table's data rows, for the editable Paid column (see "edit Paid
' on Reports" below). 2000 matches the number formats FormatReports applies.
' Row map and column map of the Reports sheet (0.10.28 layout). The FILTERS sit
' on the left in columns A:F (a column group; G is the small gap and carries the
' group's +/- control). Everything else - title, buttons, sort, Matching totals
' and the results table - starts at column H (REP_COL0), so that the filter
' boxes no longer share columns with the results table (and its hidden columns).
'   A:F rows 4-15  the filters; the addresses of the boxes did not change
'                  (B4.. and F4..), only the table moved
'   row 1-2 (H)    title and instruction text
'   row 3 (H..)    captions "Selected record:" / "Mark all as..."
'   rows 4-5 (H..) the button strip
'   row 7          Sort by / Sort direction (I7, K7)
'   rows 8/9       "Matching" labels / values
'   row 10         results header
'   row 11         first results row (the spilled formula is at H11)
' Public, because modInit and modExport read the same rows and columns and
' must never carry their own copy of these numbers.
Public Const REP_COL0 As Long = 8
' Results-table column numbers that other code keys off (counted from REP_COL0 = 1).
' 0.11.0 inserted Passes (6) after Printer and Set-up cost (13) after Consumable cost.
Public Const RC_CHARGEABLE As Long = 16
Public Const RC_NOTES As Long = 19
Public Const RC_LAST As Long = 19                ' Notes; the hidden Job ID follows it
' The Matching totals that sit above the results columns the minimum view keeps:
' Chargeable over Paper stock's neighbour, Paid and Unpaid over Chargeable / Paid.
Public Const MATCH_COL_CHARGEABLE As Long = 7
Public Const MATCH_COL_PAID As Long = 16
Public Const MATCH_COL_UNPAID As Long = 17
' The breakdown blocks to the right of the table (By print room / paper stock / department).
Public Const BRK_COL1 As Long = 22
Public Const BRK_COL2 As Long = 26
Public Const BRK_COL3 As Long = 30
Public Const REP_SORT_ROW As Long = 7
Public Const REP_MATCH_ROW As Long = 8
Public Const REP_MATCH_VAL_ROW As Long = 9
Public Const REP_HDR_ROW As Long = 10
Public Const REP_FIRST_ROW As Long = 11
Public Const REP_LBL_ROW As Long = 3
Public Const REP_BTN_ROW1 As Long = 4
Public Const REP_BTN_ROW2 As Long = 5
Private Const RESULTS_FIRST_ROW As Long = REP_FIRST_ROW
Private Const RESULTS_LAST_ROW As Long = 2000

' Reports header button layout (see DrawReportsButtons). Declared up here
' because VBA allows module-level constants only before the first procedure.
Private Const REP_BTN_TAG As String = "pcb_"
Private Const REP_BTN_GAP As Double = 4     ' between the two Paid stacks
Private Const REP_GROUP_GAP As Double = 4   ' between groups
Private Const REP_STRIP_COL As Long = REP_COL0   ' H: where the strip starts
Private Const REP_CLEAR_W As Double = 90    ' Clear all filters
Private Const REP_SEL_W As Double = 78      ' Selected record: Toggle Paid / Go to record
Private Const REP_MARK_W As Double = 68     ' Mark all as...: Paid / Unpaid
Private Const REP_EXPORT_W As Double = 140  ' Export report / Delete visible records
Private Const REP_LBL_H As Double = 13
Private Const REP_LBL_SEL As String = "lblSelected"
Private Const REP_LBL_MARK As String = "lblMarkAll"
Private Const REP_SEL_LABEL As String = "Selected record:"
Private Const REP_MARK_LABEL As String = "Mark all as..."

' A column of the consolidated range, found by its header text.
Private Function C(ByVal Header As String) As String
    C = "INDEX(" & DATA_SPILL & ",,MATCH(""" & Header & """," & DATA_HDR & ",0))"
End Function

' Column letter of the Nth column of the Reports table block (1 = first, Date/Time).
Public Function RepCol(ByVal N As Long) As String
    RepCol = Split(ThisWorkbook.Worksheets(SHEET_REPORTS).Cells(1, REP_COL0 + N - 1).Address(True, False), "$")(0)
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
    ws.Cells(SUM_TOT_ROW - 1, 1).Value = "Totals"
    ws.Cells(SUM_TOT_ROW - 1, 1).Font.Bold = True
    TotalCell ws, "B" & SUM_TOT_ROW, "Jobs", "=IFERROR(COUNTA(" & C("Job ID") & "),0)"
    TotalCell ws, "D" & SUM_TOT_ROW, "Gross", "=IFERROR(SUM(" & C("Gross Cost") & "),0)"
    TotalCell ws, "F" & SUM_TOT_ROW, "Disregarded", "=IFERROR(SUM(" & C("Disregarded") & "),0)"
    TotalCell ws, "H" & SUM_TOT_ROW, "Chargeable", "=IFERROR(SUM(" & C("Chargeable Cost") & "),0)"
    ' Snag list item 1c: the chargeable total split by paid status. Paid is
    ' "total minus paid" rather than a separate <>"Yes" SUMIFS, so the two
    ' always reconcile exactly to Chargeable by construction - a blank Paid
    ' (a job that predates the column) falls into Unpaid either way.
    TotalCell ws, "J" & SUM_TOT_ROW, "Paid", "=IFERROR(SUMIFS(" & C("Chargeable Cost") & "," & C("Paid") & ",""Yes""),0)"
    TotalCell ws, "L" & SUM_TOT_ROW, "Unpaid", "=IFERROR(SUM(" & C("Chargeable Cost") & ")-SUMIFS(" & C("Chargeable Cost") & "," & C("Paid") & ",""Yes""),0)"

    ' --- headers -----------------------------------------------------------
    ' Passes after Printer and Set-up cost after Consumable cost (0.11.0, multi-pass
    ' design decision 28): colour passes and their set-up are totals per key like the rest.
    WriteHeaderRow ws, SUM_HDR_ROW, Array("Location", "Printer", "Passes", "Paper stock", "Type", "Unit", _
                                "Jobs", "Qty", "Area m2", "Paper cost", _
                                "Consumable cost", "Set-up cost", "Gross", "Disregarded", "Chargeable")

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
    f = f & "HSTACK(lo,pr," & SumBy("Passes") & ",st," & _
        "IFNA(XLOOKUP(st,tblPapers[Description],tblPapers[Paper type]),IF(OR(" & isRoll & "," & isSheet & "),""Student supplied"",""(not in Papers)"")),"
    f = f & "IFNA(IF(XLOOKUP(st,tblPapers[Description],tblPapers[Measure])=""Sheet"",""sheets"",""metres""),IF(" & isRoll & ",""metres"",IF(" & isSheet & ",""sheets"",""-""))),"
    f = f & "COUNTIFS(" & C("Location") & ",lo," & C("Printer") & ",pr," & C("Paper Stock") & ",st),"
    f = f & SumBy("Qty") & "," & SumBy("Area m2") & "," & SumBy("Paper Cost") & ","
    f = f & SumBy("Consumable Cost") & "," & SumBy("Set-up Cost") & "," & SumBy("Gross Cost") & ","
    f = f & SumBy("Disregarded") & "," & SumBy("Chargeable Cost") & ")),"
    f = f & """No print jobs have been recorded yet."")"
    ws.Cells(SUM_FIRST_ROW, 1).Formula2 = f

    FormatSummary ws
    FormatSummaryErrors ws
    DrawLegend ws
    BuildPlanner ws
    BuildDeptBox ws
    RelockSheet ws
End Sub

' --- by department (0.10.30) -----------------------------------------------
' A small box beside the job planner (I5:M21): each department that has jobs,
' with its Jobs / Gross / Disregarded / Chargeable, and an "Everyone else" line
' for the jobs that are not a department's. It reads the Department column of
' _Data (modRegistry.WriteConsolidated), so it is pure formula - nothing to
' refresh. Placed beside, not under, the Summary spill (the spill's height is
' unknown); capped at 14 departments so it can never run into the totals row,
' with the title saying so if there are more (Reports has the full breakdown).
Private Sub BuildDeptBox(ByVal ws As Worksheet)
    Const TOP_ROW As Long = 5
    Const HDR_R As Long = 6
    Const FIRST_R As Long = 7
    Const ELSE_R As Long = 21
    Const CAP As Long = 14
    Dim dept As String, names As String, box As Range

    dept = C("Department")
    names = "UNIQUE(FILTER(" & dept & "," & dept & "<>""""))"

    Set box = ws.Range("I" & TOP_ROW & ":M" & ELSE_R)
    box.Interior.Color = RGB(238, 243, 250)
    With ws.Range("I" & TOP_ROW & ":M" & TOP_ROW)
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Range("I" & TOP_ROW).Formula2 = "=IFERROR(IF(ROWS(" & names & ")>" & CAP & _
        ",""By department (first " & CAP & " of ""&ROWS(" & names & ")&"")"",""By department""),""By department"")"

    ws.Range("I" & HDR_R).Value = "Department"
    ws.Range("J" & HDR_R).Value = "Jobs"
    ws.Range("K" & HDR_R).Value = "Gross"
    ws.Range("L" & HDR_R).Value = "Disregarded"
    ws.Range("M" & HDR_R).Value = "Chargeable"
    With ws.Range("I" & HDR_R & ":M" & HDR_R)
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(150, 160, 175)
    End With
    ws.Range("J" & HDR_R & ":M" & HDR_R).HorizontalAlignment = xlRight

    ' The names, then one spilled column per figure aligned to them.
    ws.Range("I" & FIRST_R).Formula2 = "=IFERROR(TAKE(SORT(" & names & ")," & CAP & "),"""")"
    ws.Range("J" & FIRST_R).Formula2 = "=IFERROR(IF(I" & FIRST_R & "#="""","""",COUNTIFS(" & dept & ",I" & FIRST_R & "#)),"""")"
    ws.Range("K" & FIRST_R).Formula2 = "=IFERROR(IF(I" & FIRST_R & "#="""","""",SUMIFS(" & C("Gross Cost") & "," & dept & ",I" & FIRST_R & "#)),"""")"
    ws.Range("L" & FIRST_R).Formula2 = "=IFERROR(IF(I" & FIRST_R & "#="""","""",SUMIFS(" & C("Disregarded") & "," & dept & ",I" & FIRST_R & "#)),"""")"
    ws.Range("M" & FIRST_R).Formula2 = "=IFERROR(IF(I" & FIRST_R & "#="""","""",SUMIFS(" & C("Chargeable Cost") & "," & dept & ",I" & FIRST_R & "#)),"""")"

    ws.Range("I" & ELSE_R).Value = "Everyone else"
    ws.Range("J" & ELSE_R).Formula2 = "=IFERROR(COUNTIFS(" & dept & ",""""," & C("Job ID") & ",""<>""),0)"
    ws.Range("K" & ELSE_R).Formula2 = "=IFERROR(SUMIFS(" & C("Gross Cost") & "," & dept & ",""""),0)"
    ws.Range("L" & ELSE_R).Formula2 = "=IFERROR(SUMIFS(" & C("Disregarded") & "," & dept & ",""""),0)"
    ws.Range("M" & ELSE_R).Formula2 = "=IFERROR(SUMIFS(" & C("Chargeable Cost") & "," & dept & ",""""),0)"
    With ws.Range("I" & ELSE_R & ":M" & ELSE_R)
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(150, 160, 175)
    End With

    ws.Range("J" & FIRST_R & ":J" & ELSE_R).NumberFormat = "#,##0"
    ws.Range("K" & FIRST_R & ":M" & ELSE_R).NumberFormat = CurrencyFormatCode()
    ws.Columns("I").ColumnWidth = 24
End Sub

Private Sub FormatSummary(ByVal ws As Worksheet)
    ' Warm tint (0.9.14), matching the Reports page's own totals header (row
    ' 12 there) rather than its blue results-table header - this table is
    ' itself a totals breakdown (by location/printer/paper stock), the same
    ' category as Reports' "Matching" row, not a per-job record list.
    ws.Range(ws.Cells(SUM_HDR_ROW, 1), ws.Cells(SUM_HDR_ROW, 15)).Interior.Color = RGB(244, 232, 222)
    ws.Range("C" & SUM_FIRST_ROW & ":C2000").NumberFormat = "0;-0;"               ' Passes (0 not shown)
    ws.Range("H" & SUM_FIRST_ROW & ":I2000").NumberFormat = "#,##0.00"           ' Qty, Area m2
    ws.Range("J" & SUM_FIRST_ROW & ":O2000").NumberFormat = CurrencyFormatCode() ' Paper cost .. Chargeable
    ws.Range("D" & SUM_TOT_ROW).NumberFormat = CurrencyFormatCode()
    ws.Range("F" & SUM_TOT_ROW).NumberFormat = CurrencyFormatCode()
    ws.Range("H" & SUM_TOT_ROW).NumberFormat = CurrencyFormatCode()
    ws.Range("J" & SUM_TOT_ROW).NumberFormat = CurrencyFormatCode()
    ws.Range("L" & SUM_TOT_ROW).NumberFormat = CurrencyFormatCode()
    ws.Columns("A:O").ColumnWidth = 14
    ws.Columns("A:D").ColumnWidth = 24      ' Location, Printer, Passes (the planner box shares C), Paper stock
    ws.Rows(SUM_HDR_ROW).Font.Bold = True
End Sub

' Phase 8's "error state" (design doc §11): the one place the workbook's own
' formulas already flag a genuine data-integrity break, as opposed to a
' routine per-row validation nag - a paper stock a job was costed against has
' since been renamed or removed from tblPapers, so the Type lookup in
' the first row of the table (SUM_FIRST_ROW)'s spilled formula falls back to the literal "(not in Papers)" (IFNA in
' BuildSummary above). Like Status, this is pure spilled-formula output with
' no per-cell VBA hook, so conditional formatting is the only way to colour
' it. Red reuses the exact colour the Reports disjoint-criteria warning and
' the "needs export" flag already use (modReports.BuildReports, modExport.
' RefreshExportStatus), so "error" reads the same everywhere it appears.
Private Sub FormatSummaryErrors(ByVal ws As Worksheet)
    Dim rng As Range, fc As FormatCondition
    ' The Type column: E since 0.11.0 put Passes in C.
    Set rng = ws.Range("E" & SUM_FIRST_ROW & ":E2000")
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=E" & SUM_FIRST_ROW & "=""(not in Papers)""")
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
    ' (Row 6 now holds the Department filter, 0.10.30, so the gap is gone.)
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
    s = s & "*IF($B$12="""",TRUE,IF($B$12=""Yes""," & C("Paid") & "=""Yes""," & C("Paid") & "<>""Yes""))"

    ' Chargeable cost range (0.10.25): Min cost $B$10, Max cost $B$11. Chargeable
    ' is what the student/department is billed (Gross less Disregarded), the
    ' money column Matching and the default view already use. A box that is
    ' blank, or holds something that is not a number, is ignored - the same
    ' tolerance the date boxes have. ROUND to 4 places so a value held as
    ' 12.349999999 still meets a minimum of 12.35. Both limits are inclusive.
    s = s & "*IF(ISNUMBER($B$10),ROUND(IFERROR(" & C("Chargeable Cost") & "*1,0),4)>=$B$10,TRUE)"
    s = s & "*IF(ISNUMBER($B$11),ROUND(IFERROR(" & C("Chargeable Cost") & "*1,0),4)<=$B$11,TRUE)"

    ' Has notes (0.10.25; in the More filters row, F14, since 0.10.27): Yes = the job has a note, No = it has none. A note
    ' counts when it holds something other than spaces. See NotesText for why
    ' a never-filled Notes cell has to be tested as 0 as well as "".
    s = s & "*IF($F$14=""Yes"",LEN(TRIM(" & NotesText() & "))>0,IF($F$14=""No"",LEN(TRIM(" & NotesText() & "))=0,TRUE))"

    ' Student-supplied paper (0.10.27): Yes = the paper stock is one of the two
    ' built-in Supplied stocks (not Papers rows, modCatalog.AddBuiltInStocks) or
    ' a Papers row marked Supplied by student; No = anything else, including a
    ' stock no longer in Papers. Looked up live by name in tblPapers, the same
    ' way the Summary's Type column is, so a renamed stock is read as "No".
    Dim isSup As String
    isSup = "((XLOOKUP(" & C("Paper Stock") & ",tblPapers[Description],tblPapers[Supplied by student],""No"")=""Yes"")" & _
        "+(" & C("Paper Stock") & "=""" & SUPPLIED_ROLL & """)+(" & C("Paper Stock") & "=""" & SUPPLIED_SHEET & """)>0)"
    s = s & "*IF($F$10="""",TRUE,IF($F$10=""Yes""," & isSup & ",NOT(" & isSup & ")))"
    ' Pass type (0.11.0, multi-pass design decision 32): a multi-pass job is one with
    ' colour passes, so its Passes (the _Data column) is a number; a single-pass
    ' job's is blank. Blank filter = all, like the others.
    s = s & "*IF($F$11="""",TRUE,IF($F$11=""Multi-pass"",ISNUMBER(" & C("Passes") & "),NOT(ISNUMBER(" & C("Passes") & "))))"

    ' Has a problem (0.10.27, More filters group): Yes = Status is text and not
    ' OK (the Status column holds "OK" or a joined issue list); No = Status is
    ' OK. A blank Status (never a text cell) is neither.
    Dim stt As String, isProb As String
    stt = C("Status")
    isProb = "(ISTEXT(" & stt & ")*(" & stt & "<>""OK"")*(" & stt & "<>""""))"
    s = s & "*IF($B$14="""",TRUE,IF($B$14=""Yes""," & isProb & ",ISTEXT(" & stt & ")*(" & stt & "=""OK"")))"
    ' Disregarded (0.10.27, More filters group): Paper / Consumable = that
    ' Disregard column is Yes; Both = both are.
    Dim dP As String, dC As String
    dP = "(" & C("Disregard Paper") & "=""Yes"")"
    dC = "(" & C("Disregard Consumable") & "=""Yes"")"
    s = s & "*IF($B$15="""",TRUE,IF($B$15=""Paper""," & dP & ",IF($B$15=""Consumable""," & dC & "," & dP & "*" & dC & ")))"
    ' Location (print room), added 2026-09-25 - a dropdown of registered print
    ' rooms (RefreshReportFilterLists, AllLocationCodes), same exact-match
    ' treatment as Technician/Printer/Paper Stock above. Moved from $O$4 to
    ' $F$4 (2026-09-27, see its CritCell call site) once Technician/Printer/
    ' Paper Stock shifted down a row and freed it.
    s = s & "*IF($F$4="""",TRUE," & C("Location") & "=$F$4)"
    ' Department (0.10.30): the _Data Department column, which is the Departments
    ' sheet's Name for a job whose Student Name matches it (modRegistry.
    ' WriteConsolidated) and blank otherwise. Exact match, from a dropdown.
    s = s & "*IF($B$6="""",TRUE," & C("Department") & "=$B$6)"

    Criteria = s
End Function

' The Notes column of the consolidated range as text. INDEX over a Notes cell
' that was never filled reads back as the NUMBER 0, not "" (the same phantom
' zero Paid and Student name have), so it would show as a literal "0" and
' count as a note. Notes is free text and is never a real 0, so 0 means blank.
Private Function NotesText() As String
    NotesText = "IF(" & C("Notes") & "=0,""""," & C("Notes") & "&"""")"
End Function

' The sortable columns, in results-table order (Job ID excluded - it is a
' hidden correlation column, not something offered in the sort-by list).
' Student name/no and Paid added 2026-09-22 (snag list items 2a, 2d).
Private Function ResultHeaders() As Variant
    ' Passes after Printer and Set-up cost after Consumable cost (multi-pass design
    ' decision 28). Every later column therefore sits one or two places further
    ' right than before 0.11.0; see the REP_* constants below.
    ResultHeaders = Array("Date/Time", "Location", "Student name", "Student no", _
                          "Printer", "Passes", "Paper stock", _
                          "Qty", "Unit", "Area m2", "Paper cost", _
                          "Consumable cost", "Set-up cost", "Gross", "Disregarded", "Chargeable", "Paid", _
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

    ws.Cells(1, REP_COL0).Value = "Reports"
    ws.Cells(1, REP_COL0).Font.Size = 16
    ws.Cells(1, REP_COL0).Font.Bold = True
    ws.Cells(2, REP_COL0).Value = "Find and filter print jobs across every room in this workbook. Results update as you type - " & _
        "there is no search button. Leave a box empty to ignore it. The filters can be collapsed with the - button above the column letters."

    ' The header buttons and their two small captions ("Selected record:" and
    ' "Mark all as...") are not written here. They are drawn by
    ' DrawReportsButtons (from modInit.InitialiseWorkbook, which runs after
    ' this so the row heights they are sized to are already settled) and kept
    ' in place by RepositionReportsButtons. The captions used to be cell text
    ' (O1), which stayed put when the buttons moved off a hidden column; they
    ' are now free-floating labels that move with the buttons they describe.

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
    ' Filter block layout (2026-10-02, direct user request; extended 0.10.25) -
    ' two columns, with a blank row between each group:
    '   rows 4-5   A:B  Student/dept. name, Student number
    '              D:F  Location, Technician
    '   row 6      gap
    '   rows 7-8   A:B  From date, To date
    '   rows 7-10  D:F  Printer, Paper stock, Paper type, Student-supplied paper
    '   row 9      gap (left)
    '   rows 10-12 A:B  Min cost, Max cost (chargeable), Paid - one group, no
    '              gap row between them (0.10.26, direct user request)
    '   row 13     the name/number warning (A13), directly under Paid
    '   rows 14-15 More filters, a closed nested group: Has a problem (A:B) and
    '              Has notes (D:F), then Disregarded (A:B); row 16 is its
    '              summary row (0.10.27); Sort by at row 17
    CritCell ws, "A4", "B4", "Student/dept. name"
    CritCell ws, "A5", "B5", "Student number"
    ' Department (0.10.30), directly under the student name/number pair - it is
    ' the same kind of question (who was this for). Takes row 6, which was the
    ' blank gap before the dates; the gap is gone.
    CritCell ws, "A6", "B6", "Department"
    CritCell ws, "A7", "B7", "From date"
    CritCell ws, "A8", "B8", "To date"
    ws.Range("B7:B8").NumberFormat = "dd/mm/yyyy"
    AddDateValidation ws.Range("B7"), "From date", "Leave blank for no start date."
    AddDateValidation ws.Range("B8"), "To date", "Jobs logged at any time on this date are included."

    ' Min/Max cost (0.10.25, direct user request): the chargeable cost range,
    ' inclusive at both ends, above Paid. Either box can be left blank.
    CritCell ws, "A10", "B10", "Min cost"
    CritCell ws, "A11", "B11", "Max cost"
    ws.Range("B10:B11").NumberFormat = CurrencyFormatCode()
    AddCostValidation ws.Range("B10"), "Min cost", "Show only jobs whose chargeable cost is at least this amount. Leave blank for no lower limit."
    AddCostValidation ws.Range("B11"), "Max cost", "Show only jobs whose chargeable cost is no more than this amount. Leave blank for no upper limit."

    ' Paid (Yes/No): Yes = only jobs marked Paid; No = everything not marked
    ' Paid, blank included. Left empty to ignore it.
    CritCell ws, "A12", "B12", "Paid"
    AddList ws.Range("B12"), """Yes"",""No""", "Paid", "Yes shows only paid jobs, No only unpaid ones (including jobs never marked). Leave blank to include both."

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

    ' Student-supplied paper (0.10.27, direct user request): directly under
    ' Paper type. Yes = the job's paper stock is one of the two built-in
    ' Supplied stocks or a Papers row marked Supplied by student; No = any
    ' other stock. Row 11 on this side is now empty (Has notes moved to the
    ' More filters row below).
    ' Pass type (0.11.0): row 11 on this side, directly under the paper filters.
    CritCell ws, "D11", "F11", "Pass type"
    AddList ws.Range("F11"), """Single pass"",""Multi-pass""", "Pass type", "Single pass shows ordinary jobs, Multi-pass only jobs printed in colour passes (a RISO duplicator, for example). Leave blank to include both."
    CritCell ws, "D10", "F10", "Student-supplied paper"
    AddList ws.Range("F10"), """Yes"",""No""", "Student-supplied paper", "Yes shows only jobs on paper the student supplied (Supplied (Roll), Supplied (Sheet) or a Papers row marked Supplied by student), No only jobs on stock the print room supplied. Leave blank to include both."

    ' More filters (0.10.27, direct user request): the less-used filters live
    ' in rows 14-15, a nested row group closed by default, so the main block
    ' stays short. Row 16 (formerly the gap before Sort by) is the group's
    ' summary row: Excel puts the +/- control on it (summary rows sit below
    ' their group), and it says how many of the hidden filters are set so one
    ' cannot go unnoticed.
    CritCell ws, "A14", "B14", "Has a problem"
    AddList ws.Range("B14"), """Yes"",""No""", "Has a problem", "Yes shows jobs whose Status is not OK, No only jobs whose Status is OK. Leave blank to include all."
    CritCell ws, "D14", "F14", "Has notes"
    AddList ws.Range("F14"), """Yes"",""No""", "Has notes", "Yes shows only jobs that have a note, No only jobs with none. Leave blank to include both."
    CritCell ws, "A15", "B15", "Disregarded"
    AddList ws.Range("B15"), """Paper"",""Consumable"",""Both""", "Disregarded", "Shows jobs where the paper cost, the consumable cost, or both are disregarded. Leave blank to include all."
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
    ' item 3) rather than beside the filters - row 13 is the "name and number
    ' don't match" warning, row 14 is the gap between the filters and the sort
    ' settings, and row 16 is the gap before Matching. Row 15 since 0.10.25
    ' (row 12 before the cost range and Has notes were added); the Paid filter
    ' moving up a row in 0.10.26 did not move it, the freed row went to the
    ' warning instead.
    CritCell ws, RepCol(1) & REP_SORT_ROW, RepCol(2) & REP_SORT_ROW, "Sort by"
    CritCell ws, RepCol(3) & REP_SORT_ROW, RepCol(4) & REP_SORT_ROW, "Sort direction"
    AddList ws.Range(RepCol(2) & REP_SORT_ROW), QuotedList(hdrs), "Sort by", "Which column to sort the results by."
    AddList ws.Range(RepCol(4) & REP_SORT_ROW), """Ascending"",""Descending""", "Sort direction", "Which way to sort."

    ' Export names (a Yes/No toggle that lived at N12/O12) was removed
    ' 2026-10-02, direct user request: Export report now asks Yes / No /
    ' Cancel at export time instead (modExport.AskIncludeNamesReport), the
    ' same as the room exports. The live results always show names.

    ' Spec 14.1: both criteria given, neither matching the other. Row 13
    ' (row 14 until 0.10.26, row 11 before 0.10.25, row 9 before 2026-10-02):
    ' directly under the filters, with a blank row 14 before Sort by.
    ws.Range("A13").Formula2 = "=IF(OR($B$4="""",$B$5=""""),""""," & _
        "IF(IFERROR(ROWS(FILTER(" & C("Job ID") & "," & ok & ")),0)=0," & _
        """That name and that number do not appear together on any record - check both."",""""))"
    ws.Range("A13").Font.Color = RGB(176, 0, 32)
    ws.Range("A13").Font.Bold = True

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
    ws.Range(RepCol(1) & REP_MATCH_ROW).Value = "Matching"
    ws.Range(RepCol(1) & REP_MATCH_ROW).Font.Bold = True
    MatchTotal ws, RepCol(2), "Jobs", "=IFERROR(ROWS(FILTER(" & C("Job ID") & "," & ok & ")),0)"
    MatchTotal ws, RepCol(3), "Gross", "=IFERROR(SUM(FILTER(" & C("Gross Cost") & "," & ok & ")),0)"
    MatchTotal ws, RepCol(4), "Disregarded", "=IFERROR(SUM(FILTER(" & C("Disregarded") & "," & ok & ")),0)"
    MatchTotal ws, RepCol(MATCH_COL_CHARGEABLE), "Chargeable", "=IFERROR(SUM(FILTER(" & C("Chargeable Cost") & "," & ok & ")),0)"
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
    MatchTotal ws, RepCol(MATCH_COL_PAID), "Paid", "=IFERROR(SUM(FILTER(" & C("Chargeable Cost") & "," & paidOk & ",0)),0)"
    ' The criteria are bound once (LET): this formula repeated them twice, and with
    ' the Pass type filter added that crossed Excel's 8192-character formula limit.
    MatchTotal ws, RepCol(MATCH_COL_UNPAID), "Unpaid", "=IFERROR(LET(ok0," & ok & ",cc," & C("Chargeable Cost") & _
        ",SUM(FILTER(cc,ok0,0))-SUM(FILTER(cc,ok0*(" & C("Paid") & "=""Yes""),0))),0)"
    ws.Range(RepCol(3) & REP_MATCH_VAL_ROW).NumberFormat = CurrencyFormatCode()
    ws.Range(RepCol(4) & REP_MATCH_VAL_ROW).NumberFormat = CurrencyFormatCode()
    ws.Range(RepCol(MATCH_COL_CHARGEABLE) & REP_MATCH_VAL_ROW).NumberFormat = CurrencyFormatCode()
    ws.Range(RepCol(MATCH_COL_PAID) & REP_MATCH_VAL_ROW).NumberFormat = CurrencyFormatCode()
    ws.Range(RepCol(MATCH_COL_UNPAID) & REP_MATCH_VAL_ROW).NumberFormat = CurrencyFormatCode()

    ' --- the records ------------------------------------------------------
    ' Job ID is appended after Notes and hidden - the correlation key that
    ' maps a visible row back to its source location sheet and table row for
    ' the Reports-page delete. Nothing else moves, so existing column
    ' positions are untouched.
    WriteHeaderRow ws, REP_HDR_ROW, hdrs, REP_COL0
    ws.Cells(REP_HDR_ROW, REP_COL0 + UBound(hdrs) - LBound(hdrs) + 1).Value = "Job ID"

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
    ' A job with no student name/number reads back from INDEX as the number 0, not
    ' an empty string (same as Paid, below) - shown as a literal "0" in the results
    ' and in an exported report. Student text is never a real 0, so test for it.
    rowShape = C("Job ID") & "<>" & q & q
    sName = "IF(" & rowShape & ",IF(" & C("Student Name") & "=0," & q & q & "," & C("Student Name") & ")," & q & q & ")"
    sNo = "IF(" & rowShape & ",IF(" & C("Student No") & "=0," & q & q & "," & C("Student No") & ")," & q & q & ")"

    ' A row that predates the Paid column (§5, blank Paid counts as unpaid)
    ' reads back from INDEX as the NUMBER 0, not an empty string - the
    ' classic "reference to a genuinely blank cell returns 0" Excel
    ' behaviour - so the results table would otherwise display a literal
    ' "0" for those rows. &"" does NOT fix this (0&"" is still the text
    ' "0", just stringified) - Paid only ever legitimately holds "Yes",
    ' "No" or blank, never a real 0, so testing for =0 specifically catches
    ' exactly the phantom-zero case and nothing else.
    Dim paidText As String, passesNum As String
    paidText = "IF(" & C("Paid") & "=0," & q & q & "," & C("Paid") & ")"
    ' Passes is a number or blank in _Data (a count on a multi-pass job). As a real 0
    ' for the rest, with the zero not displayed (FormatReports), it sorts as a number
    ' - blank text would sort ahead of every count in descending order.
    passesNum = "IF(ISNUMBER(" & C("Passes") & ")," & C("Passes") & ",0)"

    f = "=IFERROR(LET(" & _
        "res,FILTER(HSTACK(" & C("Date/Time") & "," & C("Location") & "," & sName & "," & sNo & _
        "," & C("Printer") & "," & passesNum & "," & C("Paper Stock") & "," & C("Qty") & "," & C("Unit") & "," & C("Area m2") & _
        "," & C("Paper Cost") & "," & C("Consumable Cost") & "," & C("Set-up Cost") & "," & C("Gross Cost") & _
        "," & C("Disregarded") & "," & C("Chargeable Cost") & "," & paidText & "," & C("Technician") & _
        "," & NotesText() & "," & C("Job ID") & ")," & ok & ",""No print jobs match those criteria.""),"
    f = f & "hdrs,{" & QuotedList(hdrs) & "},"
    f = f & "sortIdx,IFERROR(MATCH($" & RepCol(2) & "$" & REP_SORT_ROW & ",hdrs,0),0),"
    f = f & "dir,IF($" & RepCol(4) & "$" & REP_SORT_ROW & "=""Descending"",-1,1),"
    f = f & "IF(sortIdx=0,res,IFERROR(SORTBY(res,INDEX(res,,sortIdx),dir),res))"
    f = f & "),""No print jobs have been recorded yet."")"
    ws.Cells(REP_FIRST_ROW, REP_COL0).Formula2 = f
    ws.Columns(REP_COL0 + UBound(hdrs) - LBound(hdrs) + 1).Hidden = True

    ws.Range(EXPORT_SIG_CELL).Value = savedSig
    If IsDate(savedWhen) Then
        ws.Range(EXPORT_WHEN_CELL).Value = savedWhen
        ws.Range(EXPORT_WHEN_CELL).NumberFormat = "dd/mm/yyyy hh:mm"
    End If
    ws.Range(EXPORT_SIG_CELL).EntireColumn.Hidden = True

    BuildBreakdowns ws, ok
    FormatReports ws
    LockPaidColumn ws

    ' The filter labels/hints (CritCell, above) were written and wrapped
    ' before this point, while columns A/C/D/N still sat at Excel's
    ' factory-default width - FormatReports's ColumnWidth calls, just above,
    ' are the first thing to widen them. Re-fitting now, with the real widths
    ' in place, is what keeps rows 4:10 sized for the text that actually
    ' fits per line rather than for the cramped default - see CritCell's own
    ' comment for how this was found.
    ' 0.10.28: the filters now sit beside the results table, so these rows are
    ' shared with it. Fixed, uniform heights (no wrapping, no AutoFit): rows 4-5
    ' are button rows (22pt), the caption row above them is short, and the rest
    ' match the filter rows so the first table rows do not look uneven.
    ws.Range("A4:F15").WrapText = False
    ws.Range("A4:F15").VerticalAlignment = xlCenter
    ws.Rows("1:2").RowHeight = 21
    ws.Rows(REP_LBL_ROW).RowHeight = 15
    ws.Rows(REP_BTN_ROW1 & ":" & REP_BTN_ROW2).RowHeight = 22
    ws.Rows("6:15").RowHeight = 18

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

    ' Column group (0.10.28): the filter columns A:F collapse together, so a user
    ' who has already set filters can hide the controls without losing them. The
    ' +/- control lands on column G, the small gap between filters and results
    ' (summary column on the right). Grouping needs the sheet unprotected. The
    ' outline is reset first: Cells.Clear does not clear outline levels, so every
    ' rebuild would otherwise stack another level on. The row groups of earlier
    ' builds (filter block and More filters) are gone: every filter is visible.
    UnlockSheet ws
    ws.Rows("1:" & REP_HDR_ROW).Hidden = False
    ws.Rows.ClearOutline
    ws.Columns.ClearOutline
    ws.Columns("A:G").Hidden = False
    ws.Outline.SummaryColumn = xlSummaryOnRight
    ws.Columns("A:F").Group
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
    Dim c1 As String, c2 As String
    c1 = RepCol(BRK_COL1): c2 = RepCol(BRK_COL2)
    ws.Range(c1 & REP_HDR_ROW).Value = "By print room"
    ws.Range(c1 & REP_HDR_ROW).Font.Bold = True
    ws.Range(c1 & REP_FIRST_ROW).Formula2 = GroupFormula(ok, "Location")

    ws.Range(c2 & REP_HDR_ROW).Value = "By paper stock"
    ws.Range(c2 & REP_HDR_ROW).Font.Bold = True
    ws.Range(c2 & REP_FIRST_ROW).Formula2 = GroupFormula(ok, "Paper Stock")

    ' By department (0.10.30): only jobs that ARE a department's - a blank
    ' Department (an ordinary student) would otherwise show as an empty key.
    ws.Range(RepCol(BRK_COL3) & REP_HDR_ROW).Value = "By department"
    ws.Range(RepCol(BRK_COL3) & REP_HDR_ROW).Font.Bold = True
    ws.Range(RepCol(BRK_COL3) & REP_FIRST_ROW).Formula2 = GroupFormula("(" & ok & ")*(" & C("Department") & "<>"""")", "Department")

    ' Each block spills as key | Jobs | Gross | Chargeable, so the money
    ' columns are the third and fourth - Jobs is a count and must not be
    ' formatted as currency.
    ws.Range(RepCol(BRK_COL1 + 1) & REP_FIRST_ROW + 1 & ":" & RepCol(BRK_COL1 + 1) & "2000").NumberFormat = "#,##0"
    ws.Range(RepCol(BRK_COL1 + 2) & REP_FIRST_ROW + 1 & ":" & RepCol(BRK_COL1 + 3) & "2000").NumberFormat = CurrencyFormatCode()
    ws.Range(RepCol(BRK_COL2 + 1) & REP_FIRST_ROW + 1 & ":" & RepCol(BRK_COL2 + 1) & "2000").NumberFormat = "#,##0"
    ws.Range(RepCol(BRK_COL2 + 2) & REP_FIRST_ROW + 1 & ":" & RepCol(BRK_COL2 + 3) & "2000").NumberFormat = CurrencyFormatCode()
    ws.Range(RepCol(BRK_COL3 + 1) & REP_FIRST_ROW + 1 & ":" & RepCol(BRK_COL3 + 1) & "2000").NumberFormat = "#,##0"
    ws.Range(RepCol(BRK_COL3 + 2) & REP_FIRST_ROW + 1 & ":" & RepCol(BRK_COL3 + 3) & "2000").NumberFormat = CurrencyFormatCode()
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
    ws.Range(RepCol(1) & REP_MATCH_ROW & ":" & RepCol(MATCH_COL_UNPAID) & REP_MATCH_ROW).Interior.Color = RGB(244, 232, 222)
    ws.Range(RepCol(1) & REP_HDR_ROW & ":" & RepCol(RC_LAST) & REP_HDR_ROW).Interior.Color = RGB(222, 232, 244)
    ws.Range(RepCol(1) & REP_FIRST_ROW & ":" & RepCol(1) & "2000").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range(RepCol(8) & REP_FIRST_ROW & ":" & RepCol(10) & "2000").NumberFormat = "#,##0.00"      ' Qty, Unit, Area m2
    ws.Range(RepCol(11) & REP_FIRST_ROW & ":" & RepCol(RC_CHARGEABLE) & "2000").NumberFormat = CurrencyFormatCode()  ' Paper cost .. Chargeable
    ws.Range(RepCol(6) & REP_FIRST_ROW & ":" & RepCol(6) & "2000").NumberFormat = "0;-0;"        ' Passes (0 not shown)
    ' Filter block A:F (labels A and D, inputs B and F; C and E are spacers),
    ' G the small gap before the results.
    ws.Columns("A:B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 1.5
    ws.Columns("F").ColumnWidth = 22
    ws.Columns("G").ColumnWidth = 2.5
    ws.Range(RepCol(1) & ":" & RepCol(RC_LAST)).ColumnWidth = 14
    ws.Range(RepCol(2) & ":" & RepCol(5)).ColumnWidth = 22          ' Location .. Printer
    ws.Columns(RepCol(6)).ColumnWidth = 9                           ' Passes
    ws.Columns(RepCol(7)).ColumnWidth = 22                          ' Paper stock
    ws.Columns(RepCol(RC_NOTES)).ColumnWidth = 30
    ws.Columns(RepCol(1)).ColumnWidth = ColWidthForPx(180)  ' Date/Time, 180px
    ' Chargeable is wide enough to hold the Go to record / Clear all filters
    ' buttons' strip neighbours (0.10.25, DrawReportsButtons).
    ws.Columns(RepCol(RC_CHARGEABLE)).ColumnWidth = ColWidthForPx(140)
    ws.Columns(RepCol(BRK_COL1)).ColumnWidth = 22
    ws.Columns(RepCol(BRK_COL2)).ColumnWidth = 22
    ws.Columns(RepCol(BRK_COL3)).ColumnWidth = 22
    ws.Rows(REP_HDR_ROW).Font.Bold = True
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
    Set rng = ws.Cells(REP_FIRST_ROW, REP_COL0).SpillingToRange
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

    locCol = ColByHeader(ws, REP_HDR_ROW, "Location")
    jobCol = ColByHeader(ws, REP_HDR_ROW, "Job ID")
    If locCol = 0 Or jobCol = 0 Then
        Say "The Reports sheet layout looks wrong.", "The Location or Job ID column could not be found.", "Rebuild the report sheets (Refresh Locations), then try again."
        Exit Sub
    End If

    ' Broken down by room, following the same "show what you're about to
    ' lose" pattern as modJobs.RemoveRow/ClearAll.
    Set rooms = New clsDict
    For i = 1 To n
        loc = CStr(rng.Cells(i, locCol - rng.Column + 1).Value)
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
    Dim jobIds() As String, locs() As String, grp As Collection, g As Long

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
        jobIds(i) = CStr(rng.Cells(i, JobCol - rng.Column + 1).Value)
        locs(i) = CStr(rng.Cells(i, LocCol - rng.Column + 1).Value)
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
            ' A job takes its colour passes with it; whole-row deletes, last row
            ' first, keep the outline levels with their data.
            Set grp = GroupRows(lo, rowIdx)
            For g = grp.Count To 1 Step -1
                If lo.ListRows.Count = 1 Then
                    lo.ListRows(1).Delete
                Else
                    lo.ListRows(CLng(grp(g))).Range.EntireRow.Delete
                End If
            Next g
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

' ============================================================ Paid column ===
' 0.10.25: the Paid cells of the results table are READ-ONLY again, and Paid is
' changed with the "Toggle Paid" button (column N, row 2), which acts on the
' selected row or rows. 0.10.7 had made the cells directly editable with a
' Yes/No dropdown; that is gone, because it cannot be made reliable on Excel
' for Mac.
'
' Why: the results are ONE spilled formula, and anything typed into a cell of
' a spill range puts a constant there and turns the whole table into #SPILL!
' until VBA removes it again from the Change event. A trace taken on a Mac
' (Excel 16.113) showed that after a dropdown pick Excel runs NO VBA at all -
' not Change, not SelectionChange, not Calculate, not an Application.OnTime
' timer - until the user's next click, minutes later in the trace. So the
' table stayed blank until then, and no handler could prevent that. A button
' click does run its macro at once on both platforms, and nothing is ever
' typed into the spill, so it cannot be blocked.
'
' The row is read from the sheet at the moment of the click (the hidden
' Location and Job ID columns of each selected results row name the room and
' the record), so there is no selection bookkeeping to go stale.
Private Sub LockPaidColumn(ByVal ws As Worksheet)
    Dim c As Long, rng As Range
    c = ColByHeader(ws, REP_HDR_ROW, "Paid")
    If c = 0 Then Exit Sub
    Set rng = ws.Range(ws.Cells(RESULTS_FIRST_ROW, c), ws.Cells(RESULTS_LAST_ROW, c))
    rng.Validation.Delete
    rng.Locked = True
End Sub

' Sets Paid on the record behind each selected results row to the OPPOSITE of
' what the row the user is on shows: Yes becomes No, anything else (No or
' blank) becomes Yes. With several rows selected they all get that one value,
' after a confirmation. Confirmed = True skips the confirmation (for tests).
Public Sub TogglePaidSelected(Optional ByVal Confirmed As Boolean = False)
    Dim ws As Worksheet, rng As Range, ar As Range, sel As Range
    Dim locCol As Long, jobCol As Long, paidCol As Long
    Dim firstRow As Long, lastRow As Long, r As Long, i As Long, n As Long
    Dim seen As clsDict, locs() As String, jobs() As String, shown() As String
    Dim a As Variant, baseRow As Long, baseShown As String, haveBase As Boolean
    Dim newPaid As String, label As String
    Dim los As clsDict, targetWs As Worksheet, lo As ListObject, rowIdx As Long, cel As Range
    Dim changed As Long, same As Long, missing As Long, detail As String
    Dim loc As String, only As String, multi As Boolean

    Set ws = ReportsSheet()
    If ws Is Nothing Then
        Say "The Reports sheet could not be found.", "Run Refresh Locations first."
        Exit Sub
    End If

    On Error Resume Next
    Set rng = ws.Cells(REP_FIRST_ROW, REP_COL0).SpillingToRange
    On Error GoTo 0
    If rng Is Nothing Then
        Say "Select a record first.", "The Reports sheet has no results under the current filters."
        Exit Sub
    End If
    If IsError(rng.Cells(1, 1).Value2) Then
        Say "Select a record first.", "The Reports sheet has no results under the current filters."
        Exit Sub
    End If
    If Not IsNumeric(rng.Cells(1, 1).Value2) Then
        Say "There is nothing to change.", CStr(rng.Cells(1, 1).Value)
        Exit Sub
    End If

    locCol = ColByHeader(ws, REP_HDR_ROW, "Location")
    jobCol = ColByHeader(ws, REP_HDR_ROW, "Job ID")
    paidCol = ColByHeader(ws, REP_HDR_ROW, "Paid")
    If locCol = 0 Or jobCol = 0 Or paidCol = 0 Then
        Say "The Reports sheet layout looks wrong.", "The Location, Paid or Job ID column could not be found.", "Rebuild the report sheets (Refresh Locations), then try again."
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        Say "Select a record first.", "Click the Paid cell (or any cell) of the row you want to change, then press Toggle Paid."
        Exit Sub
    End If
    Set sel = Selection
    firstRow = rng.Row
    lastRow = rng.Row + rng.Rows.Count - 1

    ReDim locs(1 To 201)
    ReDim jobs(1 To 201)
    ReDim shown(1 To 201)
    Set seen = New clsDict
    For Each ar In sel.Areas
        For r = ar.Row To ar.Row + ar.Rows.Count - 1
            If r >= firstRow And r <= lastRow Then
                If Not seen.Exists(CStr(r)) Then
                    seen.Add CStr(r), 1
                    If Not IsError(ws.Cells(r, jobCol).Value2) And Not IsError(ws.Cells(r, locCol).Value2) Then
                        If Len(Trim$(CStr(ws.Cells(r, jobCol).Value2))) > 0 Then
                            n = n + 1
                            If n > 200 Then
                                Say "Too many rows selected.", "Select up to 200 records at a time.", "To mark everything a filter shows, use Mark all as..."
                                Exit Sub
                            End If
                            locs(n) = CStr(ws.Cells(r, locCol).Value2)
                            jobs(n) = Trim$(CStr(ws.Cells(r, jobCol).Value2))
                            If IsError(ws.Cells(r, paidCol).Value2) Then
                                shown(n) = ""
                            Else
                                shown(n) = Trim$(CStr(ws.Cells(r, paidCol).Value2))
                            End If
                            ' The row the cursor is on decides the new value.
                            If r = ActiveCell.Row Then
                                baseShown = shown(n)
                                haveBase = True
                            End If
                        End If
                    End If
                End If
            End If
        Next r
    Next ar

    If n = 0 Then
        Say "Select a record first.", "Click the Paid cell (or any cell) of a row that shows a print job, then press Toggle Paid."
        Exit Sub
    End If
    If Not haveBase Then baseShown = shown(1)

    If StrComp(baseShown, "Yes", vbTextCompare) = 0 Then newPaid = "No" Else newPaid = "Yes"
    label = IIf(newPaid = "Yes", "Paid", "Unpaid")
    multi = (n > 1)

    If multi And Not Confirmed Then
        If Not Ask("Mark " & n & " selected records as " & label & "?", "Toggle Paid") Then Exit Sub
    End If

    On Error GoTo Fail
    AppOff
    Set los = New clsDict
    For i = 1 To n
        loc = locs(i)
        rowIdx = 0
        Set lo = Nothing
        Set targetWs = SheetForCode(loc)
        If Not targetWs Is Nothing Then Set lo = JobsTable(targetWs)
        If Not lo Is Nothing Then rowIdx = FindReportRow(lo, jobs(i))
        If rowIdx = 0 Then
            missing = missing + 1
        Else
            Set cel = CellIn(lo, rowIdx, "Paid")
            If StrComp(Trim$(CStr(cel.Value)), newPaid, vbTextCompare) = 0 Then
                same = same + 1
            Else
                If Not los.Exists(loc) Then
                    UnlockSheet targetWs
                    los.Add loc, lo
                End If
                cel.Value = newPaid
                changed = changed + 1
                only = loc & " " & jobs(i)
            End If
        End If
    Next i
    RelockRooms los
    AppOn

    If changed > 0 Then
        If multi Then
            LogAudit "Paid edited (Reports)", "(multiple rooms)", changed & " selected record(s) marked " & label
        Else
            LogAudit "Paid edited (Reports)", locs(1), jobs(1) & ": " & IIf(Len(baseShown) = 0, "(blank)", baseShown) & " -> " & newPaid
        End If
    End If
    If missing > 0 Then
        Say missing & " record" & IIf(missing = 1, "", "s") & " could not be found.", "They may have been removed or moved since the report was drawn. Nothing was changed for them."
    End If
    Exit Sub
Fail:
    On Error Resume Next
    RelockRooms los
    On Error GoTo 0
    AppReset
    ReportError "Toggle Paid"
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
' Cell addresses match Criteria: B4 name, B5 number, B7/B8 dates, B10/B11
' min/max cost, B12 paid, F4 room, F5 technician, F7 printer, F8 paper stock,
' F9 paper type, F11 has notes. If a filter box moves, change the lists
' below AND Criteria - ClearReportFilters uses the same lists.
'
' The cost boxes count only when they hold a number (Criteria ignores anything
' else), exactly as the date boxes count only when they coerce to a date.
Public Function HasActiveFilter(ByVal ws As Worksheet) As Boolean
    Dim a As Variant
    For Each a In TextFilterCells()
        If FilterBoxHasText(ws.Range(CStr(a))) Then
            HasActiveFilter = True
            Exit Function
        End If
    Next a
    For Each a In DateFilterCells()
        If FilterBoxHasText(ws.Range(CStr(a))) Then
            If Not CBool(ws.Evaluate("ISERROR(" & CStr(a) & "*1)")) Then
                HasActiveFilter = True
                Exit Function
            End If
        End If
    Next a
    For Each a In CostFilterCells()
        If CBool(ws.Evaluate("ISNUMBER(" & CStr(a) & ")")) Then
            HasActiveFilter = True
            Exit Function
        End If
    Next a
End Function

' The filter boxes, by kind. Sort by and Sort direction are not filters and
' are in none of them.
Private Function TextFilterCells() As Variant
    TextFilterCells = Array("B4", "B5", "B12", "B14", "B15", "F4", "F5", "F7", "F8", "F9", "F10", "F11", "F14", "B6")
End Function

Private Function DateFilterCells() As Variant
    DateFilterCells = Array("B7", "B8")
End Function

Private Function CostFilterCells() As Variant
    CostFilterCells = Array("B10", "B11")
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
    Set rng = ws.Cells(REP_FIRST_ROW, REP_COL0).SpillingToRange
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

    locCol = ColByHeader(ws, REP_HDR_ROW, "Location")
    jobCol = ColByHeader(ws, REP_HDR_ROW, "Job ID")
    paidCol = ColByHeader(ws, REP_HDR_ROW, "Paid")
    If locCol = 0 Or jobCol = 0 Or paidCol = 0 Then
        Say "The Reports sheet layout looks wrong.", "The Location, Paid or Job ID column could not be found.", "Rebuild the report sheets (Refresh Locations), then try again."
        Exit Sub
    End If

    ' Per-room breakdown, as the delete prompt does, plus how many are
    ' already at the requested value (they are left alone, not rewritten).
    Set rooms = New clsDict
    For i = 1 To n
        loc = CStr(rng.Cells(i, locCol - rng.Column + 1).Value)
        If rooms.Exists(loc) Then
            rooms.Add loc, CLng(rooms.Item(loc)) + 1
        Else
            rooms.Add loc, 1
        End If
        If StrComp(Trim$(CStr(rng.Cells(i, paidCol - rng.Column + 1).Value)), NewPaid, vbTextCompare) = 0 Then same = same + 1
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
        jobIds(i) = CStr(rng.Cells(i, JobCol - rng.Column + 1).Value)
        locs(i) = CStr(rng.Cells(i, LocCol - rng.Column + 1).Value)
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
    For c = REP_COL0 To REP_COL0 + 100
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
        c = ColByHeader(ws, REP_HDR_ROW, CStr(hdrs(i)))
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

' ====================================================== filters, go to ===
' "Clear all filters" (0.10.25, direct user request): empties every filter box
' and nothing else. Sort by and Sort direction are not filters and are left
' as they are, as the Mark all / Delete visible safeguard already treats them.
' ClearContents keeps each box's dropdown and formatting. Cleared with events
' off: a change to a filter box needs no handling, and the results simply
' recalculate when calculation comes back on.
Public Sub ClearReportFilters(ByVal ws As Worksheet)
    Dim a As Variant

    On Error GoTo Fail
    AppOff
    For Each a In TextFilterCells()
        ws.Range(CStr(a)).ClearContents
    Next a
    For Each a In DateFilterCells()
        ws.Range(CStr(a)).ClearContents
    Next a
    For Each a In CostFilterCells()
        ws.Range(CStr(a)).ClearContents
    Next a
    AppOn
    Exit Sub
Fail:
    AppReset
    ReportError "Clear all filters"
End Sub

' "Go to record" (0.10.25, direct user request): takes the selected row of the
' results to the record itself - its print room's sheet, with the whole row of
' the job table selected and scrolled into view if it is not already. The row
' is found the way Delete visible and Mark all find theirs: the hidden
' Location and Job ID columns of the results row name the sheet and the record.
Public Sub GoToReportRecord()
    Dim ws As Worksheet, r As Long, a As Variant
    Dim locCol As Long, jobCol As Long, loc As String, job As String
    Dim target As Worksheet, lo As ListObject, idx As Long, rowRng As Range

    Set ws = ReportsSheet()
    If ws Is Nothing Then
        Say "The Reports sheet could not be found.", "Run Refresh Locations first."
        Exit Sub
    End If

    r = ActiveCell.Row
    If r < RESULTS_FIRST_ROW Or r > RESULTS_LAST_ROW Then
        Say "Select a record first.", "Click any cell in the row of the record you want to see, then press Go to record."
        Exit Sub
    End If

    ' A real record on this row: Date/Time, column A, is a number. (Empty past
    ' the results, text for FILTER's "no jobs match" message, an error while a
    ' typed constant is blocking the spill.)
    a = ws.Cells(r, REP_COL0).Value2
    If IsError(a) Or IsEmpty(a) Then
        Say "That row is not a record.", "Select a row that shows a print job, then press Go to record."
        Exit Sub
    End If
    If Not IsNumeric(a) Then
        Say "That row is not a record.", CStr(a)
        Exit Sub
    End If

    locCol = ColByHeader(ws, REP_HDR_ROW, "Location")
    jobCol = ColByHeader(ws, REP_HDR_ROW, "Job ID")
    If locCol = 0 Or jobCol = 0 Then
        Say "The Reports sheet layout looks wrong.", "The Location or Job ID column could not be found.", "Rebuild the report sheets (Refresh Locations), then try again."
        Exit Sub
    End If
    If IsError(ws.Cells(r, locCol).Value2) Or IsError(ws.Cells(r, jobCol).Value2) Then
        Say "That row is not a record.", "Select a row that shows a print job, then press Go to record."
        Exit Sub
    End If
    loc = CStr(ws.Cells(r, locCol).Value2)
    job = Trim$(CStr(ws.Cells(r, jobCol).Value2))

    Set target = SheetForCode(loc)
    If target Is Nothing Then
        Say "That print room could not be found.", "The room '" & loc & "' is not in this workbook any more.", "Run Refresh Locations, then try again."
        Exit Sub
    End If
    Set lo = JobsTable(target)
    If Not lo Is Nothing Then idx = FindReportRow(lo, job)
    If idx = 0 Then
        Say "That record could not be found.", "It may have been removed from " & LocValue(target, "LOC_Name") & " since the report was drawn.", "Look at the row again, then retry."
        Exit Sub
    End If

    On Error GoTo Fail
    If target.Visible <> xlSheetVisible Then target.Visible = xlSheetVisible
    target.Activate
    Set rowRng = lo.ListRows(idx).Range
    ScrollRowIntoView rowRng.Cells(1, 1)
    rowRng.Select
    rowRng.Cells(1, 1).Activate
    Exit Sub
Fail:
    ReportError "Go to record"
End Sub

' Scrolls the window only when the cell is not already on screen, putting it a
' few rows down from the top so the rows above it are still partly in view.
Private Sub ScrollRowIntoView(ByVal cel As Range)
    Dim r As Long
    On Error Resume Next
    If Intersect(ActiveWindow.VisibleRange, cel) Is Nothing Then
        r = cel.Row - 3
        If r < 1 Then r = 1
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = r
    End If
    On Error GoTo 0
End Sub

' ===================================================== header buttons ===
' The Reports header strip (rows 3-5, from column H, 0.10.28 layout; the diagram below shows the order, not the old rows). Seven buttons and two
' small captions in four blocks, left to right:
'
'   row 1  [Clear all filters]  Selected record:   Mark all as...  [Export report...]
'   row 2                       [Toggle Paid]      [Paid]
'   row 3                       [Go to record]     [Unpaid]        [Delete visible records...]
'
' The Paid controls sit together, and the captions say what each acts on:
' Toggle Paid works on the record(s) selected in the results (as does Go to
' record, which is why they share a block); Mark all as... works on every
' record the filters show. Row 2 is left empty at both ends so the strip stays
' clear of A2's instruction text, which runs to ~689pt (the first block starts
' at N, ~639pt on the test machine, and its only buttons are on rows 1 and
' 3; the first row-2 button starts ~94pt further right, at ~733pt).
'
' Nothing here depends on cell contents. The buttons and both captions are
' free-floating form controls, and RepositionReportsButtons lays the whole
' strip out from one anchor - the first visible column from N - packing the
' blocks one after another by fixed widths and gaps, so a column hide or width
' change moves the strip as one piece and the captions can never be left
' behind. (The "Mark all as..." caption used to be cell text at O1, which
' stayed put when its buttons moved.) The captions are Forms labels rather
' than text boxes: a label runs no macro and is not selected by a click, so it
' needs no OnAction for HealButtons to repair after a file rename and cannot
' replace the cell selection Toggle Paid reads.
'
' The strip fits the first screenful (every Left <= 900pt, test-reports.ps1):
' with the default columns the blocks start at ~639, 733, 815 and 887pt (the
' first run of 0.10.26 measured N at 639pt, not the 612pt first estimated, and
' failed at 915pt). The widths are in the REP_*_W constants above; widening
' one needs the others checked against that limit.
'
' The buttons are free-floating, so they do not follow column changes by
' themselves, and Excel raises no event for a hide or a width change. So
' RepositionReportsButtons re-lays them out from the columns as they are now.
' It runs after the build, on every Reports activation and on every selection
' change there (ThisWorkbook), and only touches a shape that is actually out of
' place.
'
' Every Reports button name must stay within 31 characters and hold its macro
' ("pcb_<macro>_<row>_<col>"): modRegistry.HealButtons falls back on the name
' to recover a macro. Keep the macro names short - that is what the shape
' name tag "DelVis" used to paper over, and what broke the Delete button once
' the file was renamed. (The row and column in the name are only a label; they
' do not say where the button sits.)

Public Sub DrawReportsButtons(ByVal ws As Worksheet)
    DrawOne ws, REP_BTN_ROW1, REP_STRIP_COL, "Clear all filters", "btnClearFilters", REP_CLEAR_W
    DrawOne ws, REP_BTN_ROW1, REP_STRIP_COL + 1, "Toggle Paid", "btnTogglePaid", REP_SEL_W
    DrawOne ws, REP_BTN_ROW2, REP_STRIP_COL + 1, "Go to record", "btnGoToRecord", REP_SEL_W
    DrawOne ws, REP_BTN_ROW1, REP_STRIP_COL + 2, "Paid", "btnMarkPaid", REP_MARK_W
    DrawOne ws, REP_BTN_ROW2, REP_STRIP_COL + 2, "Unpaid", "btnMarkUnpaid", REP_MARK_W
    DrawOne ws, REP_BTN_ROW1, REP_STRIP_COL + 6, "Export report...", "btnExportReport", REP_EXPORT_W
    DrawOne ws, REP_BTN_ROW2, REP_STRIP_COL + 6, "Delete visible records...", "btnDeleteVisible", REP_EXPORT_W
    DrawReportLabel ws, REP_LBL_SEL, REP_SEL_LABEL, REP_SEL_W
    DrawReportLabel ws, REP_LBL_MARK, REP_MARK_LABEL, REP_MARK_W
    RepositionReportsButtons ws
End Sub

Public Sub RepositionReportsButtons(ByVal ws As Worksheet)
    Dim c As Long, clearLeft As Double, selLeft As Double
    Dim markLeft As Double, expLeft As Double

    c = VisibleColumnFrom(ws, REP_STRIP_COL, 0)
    If c = 0 Then Exit Sub
    clearLeft = ws.Cells(1, c).Left
    selLeft = clearLeft + REP_CLEAR_W + REP_GROUP_GAP
    markLeft = selLeft + REP_SEL_W + REP_BTN_GAP
    expLeft = markLeft + REP_MARK_W + REP_GROUP_GAP

    PlaceReportButton ws, "btnClearFilters", clearLeft, ws.Rows(REP_BTN_ROW1).Top + 0.5, REP_CLEAR_W, ws.Rows(REP_BTN_ROW1).Height - 1
    PlaceReportButton ws, "btnTogglePaid", selLeft, ws.Rows(REP_BTN_ROW1).Top + 0.5, REP_SEL_W, ws.Rows(REP_BTN_ROW1).Height - 1
    PlaceReportButton ws, "btnGoToRecord", selLeft, ws.Rows(REP_BTN_ROW2).Top + 0.5, REP_SEL_W, ws.Rows(REP_BTN_ROW2).Height - 1
    PlaceReportButton ws, "btnMarkPaid", markLeft, ws.Rows(REP_BTN_ROW1).Top + 0.5, REP_MARK_W, ws.Rows(REP_BTN_ROW1).Height - 1
    PlaceReportButton ws, "btnMarkUnpaid", markLeft, ws.Rows(REP_BTN_ROW2).Top + 0.5, REP_MARK_W, ws.Rows(REP_BTN_ROW2).Height - 1
    PlaceReportButton ws, "btnExportReport", expLeft, ws.Rows(REP_BTN_ROW1).Top + 0.5, REP_EXPORT_W, ws.Rows(REP_BTN_ROW1).Height - 1
    PlaceReportButton ws, "btnDeleteVisible", expLeft, ws.Rows(REP_BTN_ROW2).Top + 0.5, REP_EXPORT_W, ws.Rows(REP_BTN_ROW2).Height - 1
    PlaceReportLabel ws, REP_LBL_SEL, selLeft, REP_SEL_W
    PlaceReportLabel ws, REP_LBL_MARK, markLeft, REP_MARK_W
End Sub

' The first column from StartCol on that is not hidden and whose left edge is
' at or past MinLeft. 0 if there is none within reach.
Private Function VisibleColumnFrom(ByVal ws As Worksheet, ByVal StartCol As Long, ByVal MinLeft As Double) As Long
    Dim c As Long
    For c = StartCol To StartCol + 40
        If Not ws.Columns(c).Hidden Then
            If ws.Cells(1, c).Left >= MinLeft - 0.5 Then
                VisibleColumnFrom = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function ReportButton(ByVal ws As Worksheet, ByVal Macro As String) As Button
    Dim i As Long, prefix As String
    prefix = REP_BTN_TAG & Macro & "_"
    For i = 1 To ws.Buttons.Count
        If Left$(ws.Buttons(i).Name, Len(prefix)) = prefix Then
            Set ReportButton = ws.Buttons(i)
            Exit Function
        End If
    Next i
End Function

Private Sub PlaceReportButton(ByVal ws As Worksheet, ByVal Macro As String, ByVal L As Double, _
                              ByVal T As Double, ByVal W As Double, ByVal H As Double)
    Dim b As Button
    Set b = ReportButton(ws, Macro)
    If b Is Nothing Then Exit Sub
    ' xlMove (0.10.28), not free-floating: Excel raises no event when the filter
    ' column group is collapsed or expanded, so a free-floating strip would stay
    ' where it was until the next click. xlMove makes the shape follow its cell
    ' when columns to its left appear or disappear, without ever resizing it
    ' (xlMoveAndSize is what resized buttons before). The re-pack below still
    ' runs on every activation and selection change for width changes.
    If b.Placement <> xlMove Then b.Placement = xlMove
    If Abs(b.Left - L) > 0.5 Then b.Left = L
    If Abs(b.Top - T) > 0.5 Then b.Top = T
    If Abs(b.Width - W) > 0.5 Then b.Width = W
    If Abs(b.Height - H) > 0.5 Then b.Height = H
End Sub

' A caption above a block of buttons: a Forms label, free-floating like the
' buttons, bottom-aligned in row 1 so it sits directly over the block's first
' button (row 2). Any earlier label of the same name is deleted first, so a
' rebuild never leaves two. 8pt bold, which fits both captions inside their
' block's width on one line - a label wraps rather than overflows, and a
' wrapped caption would lose its second line to the fixed height.
Private Sub DrawReportLabel(ByVal ws As Worksheet, ByVal Nm As String, ByVal Caption As String, ByVal W As Double)
    Dim lbl As Excel.Label, i As Long
    For i = ws.Labels.Count To 1 Step -1
        If StrComp(ws.Labels(i).Name, REP_BTN_TAG & Nm, vbBinaryCompare) = 0 Then ws.Labels(i).Delete
    Next i
    Set lbl = ws.Labels.Add(ws.Cells(1, REP_STRIP_COL).Left, ws.Rows(REP_LBL_ROW).Top, W, REP_LBL_H)
    lbl.Name = REP_BTN_TAG & Nm
    lbl.Placement = xlFreeFloating ' see modInit.DrawOne's comment
    lbl.Caption = Caption
    ' Cosmetic only: a failure here must not stop the build.
    On Error Resume Next
    lbl.Characters.Font.Size = 8
    lbl.Characters.Font.Bold = True
    On Error GoTo 0
End Sub

' Moves a caption to L, at the bottom of row 1. Only touches a property that is
' actually out of place, so the call is cheap on every selection change.
Private Sub PlaceReportLabel(ByVal ws As Worksheet, ByVal Nm As String, ByVal L As Double, ByVal W As Double)
    Dim lbl As Excel.Label, i As Long, T As Double
    For i = 1 To ws.Labels.Count
        If StrComp(ws.Labels(i).Name, REP_BTN_TAG & Nm, vbBinaryCompare) = 0 Then
            Set lbl = ws.Labels(i)
            Exit For
        End If
    Next i
    If lbl Is Nothing Then Exit Sub
    T = ws.Rows(REP_LBL_ROW).Top + ws.Rows(REP_LBL_ROW).Height - REP_LBL_H - 0.5
    If lbl.Placement <> xlMove Then lbl.Placement = xlMove
    If Abs(lbl.Left - L) > 0.5 Then lbl.Left = L
    If Abs(lbl.Top - T) > 0.5 Then lbl.Top = T
    If Abs(lbl.Width - W) > 0.5 Then lbl.Width = W
    If Abs(lbl.Height - REP_LBL_H) > 0.5 Then lbl.Height = REP_LBL_H
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
    ws.Range("A" & SUM_FIRST_ROW & ":BZ5000").ClearContents
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
    ws.Range(ColLetter & REP_MATCH_ROW).Value = Label
    ws.Range(ColLetter & REP_MATCH_ROW).Font.Bold = True
    ws.Range(ColLetter & REP_MATCH_VAL_ROW).Formula2 = f
    ws.Range(ColLetter & REP_MATCH_VAL_ROW).Font.Bold = True
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

Private Sub WriteHeaderRow(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal Headers As Variant, Optional ByVal FirstCol As Long = 1)
    Dim i As Long
    For i = LBound(Headers) To UBound(Headers)
        ws.Cells(RowNo, FirstCol + i - LBound(Headers)).Value = Headers(i)
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
    Set rng = ws.Cells(REP_FIRST_ROW, REP_COL0).SpillingToRange
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

' A cost box (0.10.25): a number of 0 or more, or blank. Like the date rule it
' only fires on manual entry, and Criteria ignores anything that is not a
' number, so a value pasted past it simply does not filter.
Private Sub AddCostValidation(ByVal target As Range, ByVal Title As String, ByVal Msg As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
        .IgnoreBlank = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = "Enter an amount of 0 or more, or leave blank."
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

    ' Department (0.10.30): every department on the Departments sheet, active
    ' or not - an old job for a deactivated one must still be findable.
    ApplyTo ws, ws.Range("B6"), AllDepartments(), "REP|Department", _
        "Department", "Choose a department to see only its jobs, or leave blank to include all."
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
