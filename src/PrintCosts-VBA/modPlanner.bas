Attribute VB_Name = "modPlanner"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' The job planner: a boxed utility at the top of the Summary sheet that lets
' someone price a print job BEFORE it is committed, then (optionally) add it,
' prefilled, to a print room's job table.
'
' It is a utility, separate from the Summary report below it. It reads the
' CURRENT catalogue prices (the same ones a new job would be stamped with -
' modSnapshot.StampRow) and does the arithmetic of the job-row formulas
' (Area m2 / Paper Cost / Consumable Cost / Gross Cost, shipped in the .xlsx):
'
'   Paper cost      = ROUND(Qty x unit cost, SET_ROUND_DP)           Qty in metres or sheets
'   Area m2         = roll : print width / 1000 x metres
'                     sheet: width / 1000 x height / 1000 x sheets
'   Consumable cost = ROUND(Area m2 x printer rate per m2, SET_ROUND_DP)
'   Overall         = paper + consumable            (gross: disregard flags are not applied)
'
' The roll length is always entered in METRES here, whatever Roll length unit a
' print room uses; AddPlannedJob converts it for a Centimetres room.
'
' The figures are written by VBA when an input changes (Workbook_SheetChange ->
' PlannerChanged) and when the Summary sheet is activated (so a price edited on
' Papers/Printers is picked up). They are not stored anywhere else and are never
' read back by anything except AddPlannedJob, which recomputes first.

' ---- layout (Summary sheet). Tests and modReports read these. -------------
Public Const PLN_TOP_ROW As Long = 5
Public Const PLN_BOTTOM_ROW As Long = 21
Public Const PLN_LAST_COL As Long = 6              ' box is A:F
Public Const PLN_PRINTER As String = "B7"
Public Const PLN_PAPER As String = "B8"
Public Const PLN_QTY As String = "B9"
Public Const PLN_WIDTH As String = "B10"
Public Const PLN_SIZE As String = "B11"
Public Const PLN_PAPER_COST As String = "B14"
Public Const PLN_INK_COST As String = "B15"
Public Const PLN_TOTAL_COST As String = "B16"
Public Const PLN_MESSAGE As String = "A17"
Public Const PLN_ROOM As String = "B19"
Public Const PLN_ADD_ROW As Long = 19
Public Const PLN_ADD_COL As Long = 3

Private Const R_QTY_LABEL As String = "A9"
Private Const TAG_PRN As String = "PLN|PRN"
Private Const TAG_STK As String = "PLN|STK"
Private Const TAG_LOC As String = "PLN|LOC"

' The outcome of reading the box. A Type must sit in the declarations section,
' above the first procedure.
Private Type PlanResult
    Message As String      ' what is missing / wrong; "" when an estimate was produced
    IsError As Boolean     ' True = a conflict (shown red); False = just not filled in yet
    PaperCost As Double
    InkCost As Double
    Qty As Double
    WidthMM As Double      ' 0 = not entered
    SizeName As String
    Printer As String
    Stock As String
End Type

Private Const NAVY As Long = 6567967      ' RGB(31, 56, 100)
Private Const BOX_FILL As Long = 16446446 ' RGB(238, 243, 250)

' ================================================================= build ===
' Draws the box. Called by modReports.BuildSummary after the rest of the sheet
' (so the column widths are final) and before it relocks the sheet.
Public Sub BuildPlanner(ByVal ws As Worksheet)
    Dim box As Range, r As Long

    Set box = ws.Range(ws.Cells(PLN_TOP_ROW, 1), ws.Cells(PLN_BOTTOM_ROW, PLN_LAST_COL))
    box.Interior.Color = BOX_FILL
    box.Font.Color = RGB(0, 0, 0)

    With ws.Range(ws.Cells(PLN_TOP_ROW, 1), ws.Cells(PLN_TOP_ROW, PLN_LAST_COL))
        .Interior.Color = NAVY
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Cells(PLN_TOP_ROW, 1).Value = "Plan a print job"

    ws.Cells(PLN_TOP_ROW + 1, 1).Value = "Choose the options to see the expected cost before committing to a job. " & _
        "Nothing is recorded until you press Add to print room."
    ws.Cells(PLN_TOP_ROW + 1, 1).Font.Italic = True
    ws.Cells(PLN_TOP_ROW + 1, 1).Font.Color = RGB(110, 110, 110)

    ws.Range("A7").Value = "Printer"
    ws.Range("A8").Value = "Paper"
    ws.Range(R_QTY_LABEL).Value = "Roll length (metres)"
    ws.Range("A10").Value = "Print width mm"
    ws.Range("A11").Value = "Sheet size"
    ws.Range("A7:A11").Font.Bold = True

    StyleBox ws.Range(PLN_PRINTER)
    StyleBox ws.Range(PLN_PAPER)
    StyleBox ws.Range(PLN_QTY)
    StyleBox ws.Range(PLN_WIDTH)
    StyleBox ws.Range(PLN_SIZE)
    ws.Range(PLN_QTY).NumberFormat = "General"
    ws.Range(PLN_WIDTH).NumberFormat = "General"
    ws.Range("B7:B11").HorizontalAlignment = xlLeft
    ws.Range("C7:C11").Font.Italic = True
    ws.Range("C7:C11").Font.Color = RGB(110, 110, 110)

    ws.Range("A13").Value = "Estimated cost"
    ws.Range("A13:F13").Font.Bold = True
    ws.Range("A13:F13").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("A13:F13").Borders(xlEdgeBottom).Color = RGB(150, 160, 175)
    ws.Range("A14").Value = "Paper"
    ws.Range("A15").Value = "Ink (consumable)"
    ws.Range("A16").Value = "Overall"
    ws.Range("A16:B16").Font.Bold = True
    With ws.Range("B14:B16")
        .Interior.Color = RGB(217, 217, 217)   ' calculated automatically (legend)
        .Font.Italic = True
        .NumberFormat = CurrencyFormatCode()
        .HorizontalAlignment = xlRight
        .Locked = True
    End With
    ws.Range(PLN_TOTAL_COST).Font.Bold = True

    ws.Range("A19").Value = "Add to print room"
    ws.Range("A19").Font.Bold = True
    StyleBox ws.Range(PLN_ROOM)
    ws.Cells(PLN_BOTTOM_ROW - 1, 1).Value = "Student, technician and the rest are completed on the room's own sheet."
    ws.Cells(PLN_BOTTOM_ROW - 1, 1).Font.Italic = True
    ws.Cells(PLN_BOTTOM_ROW - 1, 1).Font.Color = RGB(110, 110, 110)

    ' The box: a medium outline all round, nothing inside.
    With box
        .BorderAround xlContinuous, xlMedium, , RGB(31, 56, 100)
    End With

    ' Row heights are left to the sheet's default; the button sits in row 19.
    For r = PLN_TOP_ROW To PLN_BOTTOM_ROW
        ws.Rows(r).RowHeight = IIf(r = PLN_ADD_ROW, 24, 17)
    Next r
    ws.Rows(PLN_TOP_ROW).RowHeight = 22

    BindPlannerLists ws
    PlannerRecalc ws
End Sub

Private Sub StyleBox(ByVal target As Range)
    target.Locked = False
    target.Interior.Color = RGB(255, 255, 255)
    target.Borders(xlEdgeLeft).LineStyle = xlContinuous
    target.Borders(xlEdgeLeft).Color = RGB(46, 100, 168)
    target.Borders(xlEdgeLeft).Weight = xlMedium
End Sub

' ================================================================= lists ===
' Every active printer permitted at ANY print room, sorted.
Private Function AllRoomPrinters() As Collection
    Dim out As Collection, seen As clsDict, v As Variant, m As Variant
    Set out = New Collection
    Set seen = New clsDict
    For Each v In LocationSheets()
        For Each m In PrintersFor(v)
            ' A multi-pass printer is not offered: the planner prices one flat rate
            ' and cannot price a job's colour passes (multi-pass design, ARCHITECTURE 17.2).
            If Not seen.Exists(CStr(m)) And Not Prn(CStr(m)).MultiPass Then
                seen.Add CStr(m), True
                out.Add CStr(m)
            End If
        Next m
    Next v
    Set AllRoomPrinters = SortedTextCollection(out)
End Function

' Every active stock that some printer offered above can take, sorted.
Private Function AllRoomStocks(ByVal Printers As Collection) As Collection
    Dim out As Collection, seen As clsDict, m As Variant, s As Collection, i As Long
    Set out = New Collection
    Set seen = New clsDict
    For Each m In Printers
        Set s = StocksFor(CStr(m))
        For i = 1 To s.Count
            If Not seen.Exists(CStr(s(i))) Then
                seen.Add CStr(s(i)), True
                out.Add CStr(s(i))
            End If
        Next i
    Next m
    Set AllRoomStocks = SortedTextCollection(out)
End Function

' Print rooms (by name) whose permitted printers include Model. A blank Model
' lists every room.
Private Function RoomsFor(ByVal Model As String) As Collection
    Dim out As Collection, v As Variant, nm As String
    Set out = New Collection
    For Each v In LocationSheets()
        nm = LocValue(v, "LOC_Name")
        If Len(nm) = 0 Then nm = v.Name
        If Len(Model) = 0 Then
            out.Add nm
        ElseIf InList(LocValue(v, "LOC_Printers"), Model) Then
            out.Add nm
        End If
    Next v
    Set RoomsFor = SortedTextCollection(out)
End Function

' Same convention as the job table: nothing is removed from the lists, items the
' other field rules out are suffixed (unavailable) and remain pickable (the
' estimate then explains why it cannot be given).
Public Sub BindPlannerLists(ByVal ws As Worksheet)
    Dim prn As String, stk As String, printers As Collection

    prn = CleanPick(CStr(ws.Range(PLN_PRINTER).Value))
    stk = CleanPick(CStr(ws.Range(PLN_PAPER).Value))
    Set printers = AllRoomPrinters()

    ApplyTo ws, ws.Range(PLN_PRINTER), MarkPrinters(printers, stk), TAG_PRN, "Printer", _
        "Every printer available in any print room. Ones that cannot take the chosen paper are marked (unavailable)."
    ApplyTo ws, ws.Range(PLN_PAPER), MarkStocks(AllRoomStocks(printers), prn), TAG_STK, "Paper", _
        "Every paper available in any print room. Ones the chosen printer cannot take are marked (unavailable)."
    BindRoomList ws, prn

    ' ApplyTo relocks the sheet each time it finishes, so open it again for the
    ' rules below (a validation change on a protected sheet is a 1004).
    UnlockSheet ws
    With ws.Range(PLN_SIZE).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=RNG_STD_SIZES"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Sheet size"
        .InputMessage = "Only for 'Supplied (Sheet)': the nearest standard size to the sheet the student brought."
        .ErrorTitle = "Sheet size"
        .ErrorMessage = "Choose one of the standard sizes listed on Settings."
    End With
    AddDecimalRule ws.Range(PLN_QTY), "Amount", "Metres of roll, or number of sheets. Must be greater than zero."
    AddDecimalRule ws.Range(PLN_WIDTH), "Print width", "Optional for catalogue rolls (blank = full roll width). Required for 'Supplied (Roll)'. In millimetres."

    RelockSheet ws
End Sub

Private Sub BindRoomList(ByVal ws As Worksheet, ByVal Model As String)
    ApplyTo ws, ws.Range(PLN_ROOM), RoomsFor(Model), TAG_LOC, "Add to print room", _
        "The print room to add this job to. Only rooms that have the chosen printer are listed."
End Sub

Private Sub AddDecimalRule(ByVal target As Range, ByVal Title As String, ByVal Msg As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
        .IgnoreBlank = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = Msg
    End With
End Sub

' ============================================================= estimate ===
Private Function PlanFor(ByVal ws As Worksheet) As PlanResult
    Dim res As PlanResult, p As clsPrinterDef, s As clsStock
    Dim w As Double, h As Double, area As Double, dp As Long, usedW As Double
    Dim cap As String

    res.Printer = CleanPick(CStr(ws.Range(PLN_PRINTER).Value))
    res.Stock = CleanPick(CStr(ws.Range(PLN_PAPER).Value))
    res.Qty = NumOf(ws.Range(PLN_QTY))
    res.WidthMM = NumOf(ws.Range(PLN_WIDTH))
    res.SizeName = Trim$(CStr(ws.Range(PLN_SIZE).Value))
    dp = RoundDP()

    If Len(res.Printer) = 0 Then
        res.Message = "Choose a printer."
        PlanFor = res: Exit Function
    End If
    Set p = Prn(res.Printer)
    If Not p.Found Then
        res.Message = "'" & res.Printer & "' is not on the Printers sheet."
        res.IsError = True
        PlanFor = res: Exit Function
    End If
    If Len(res.Stock) = 0 Then
        res.Message = "Choose a paper."
        PlanFor = res: Exit Function
    End If
    Set s = Stock(res.Stock)
    If Not s.Found Then
        res.Message = "'" & res.Stock & "' is not on the Papers sheet."
        res.IsError = True
        PlanFor = res: Exit Function
    End If
    If Not Compatible(p.Model, s.Description) Then
        res.Message = p.Model & " cannot take " & s.Description & " - it takes " & CapacityText(p) & "."
        res.IsError = True
        PlanFor = res: Exit Function
    End If

    If res.Qty <= 0 Then
        If s.Measure = "Sheet" Then
            res.Message = "Enter the number of sheets."
        Else
            res.Message = "Enter the roll length in metres."
        End If
        PlanFor = res: Exit Function
    End If

    If s.Measure = "Sheet" Then
        w = s.WidthMM: h = s.HeightMM
        If s.PerJobSize Then
            If Len(res.SizeName) = 0 Then
                res.Message = "Choose the sheet size the student is supplying."
                PlanFor = res: Exit Function
            End If
            If Not StdSizeDims(res.SizeName, w, h) Then
                res.Message = "'" & res.SizeName & "' is not a standard size on Settings."
                res.IsError = True
                PlanFor = res: Exit Function
            End If
            If Not FitsWithinMaxSheet(w, h, p.MaxSheetWidthMM, p.MaxSheetHeightMM) Then
                res.Message = res.SizeName & " is larger than " & p.Model & " can take (" & CapacityText(p) & ")."
                res.IsError = True
                PlanFor = res: Exit Function
            End If
        End If
        area = (w / 1000#) * (h / 1000#) * res.Qty
    Else
        usedW = res.WidthMM
        If s.PerJobSize Then
            If usedW <= 0 Then
                res.Message = "Enter the print width of the roll the student is supplying."
                PlanFor = res: Exit Function
            End If
            If usedW > p.MaxRollWidthMM Then
                res.Message = "A " & Mm(usedW) & " mm width is wider than " & p.Model & " can take (" & CapacityText(p) & ")."
                res.IsError = True
                PlanFor = res: Exit Function
            End If
        Else
            If usedW <= 0 Then usedW = s.WidthMM
            If usedW > s.WidthMM Then
                res.Message = "The print width cannot exceed the " & Mm(s.WidthMM) & " mm roll."
                res.IsError = True
                PlanFor = res: Exit Function
            End If
        End If
        area = (usedW / 1000#) * res.Qty
    End If

    res.PaperCost = Application.WorksheetFunction.Round(res.Qty * s.Cost, dp)
    res.InkCost = Application.WorksheetFunction.Round(area * p.RatePerM2, dp)
    PlanFor = res
End Function

' Rewrites the labels, hints, grey-outs and estimate from the current inputs.
Public Sub PlannerRecalc(ByVal ws As Worksheet)
    Dim res As PlanResult, p As clsPrinterDef, s As clsStock, isSheet As Boolean
    Dim ev As Boolean, perJob As Boolean, symbol As String

    ev = Application.EnableEvents
    Application.EnableEvents = False
    UnlockSheet ws

    res = PlanFor(ws)
    Set p = Prn(res.Printer)
    Set s = Stock(res.Stock)
    isSheet = (s.Found And s.Measure = "Sheet")
    perJob = (s.Found And s.PerJobSize)
    symbol = CurrencySymbol()

    ' --- labels and hints -------------------------------------------------
    ws.Range(R_QTY_LABEL).Value = IIf(isSheet, "Number of sheets", "Roll length (metres)")

    If p.Found Then
        ws.Range("C7").Value = "Takes " & CapacityText(p)
    ElseIf Len(res.Printer) > 0 Then
        ws.Range("C7").Value = "Not on the Printers sheet"
    Else
        ws.Range("C7").Value = ""
    End If

    If s.Found Then
        If perJob Then
            ws.Range("C8").Value = "Student supplied - no paper charge"
        ElseIf isSheet Then
            ws.Range("C8").Value = "Sheet " & Mm(s.WidthMM) & " x " & Mm(s.HeightMM) & " mm, " & _
                symbol & Format$(s.Cost, "#,##0.00") & " per sheet"
        Else
            ws.Range("C8").Value = "Roll " & Mm(s.WidthMM) & " mm wide, " & _
                symbol & Format$(s.Cost, "#,##0.00") & " per metre"
        End If
    Else
        ws.Range("C8").Value = ""
    End If

    If isSheet Then
        ws.Range("C9").Value = "Sheets to print"
        ws.Range("C10").Value = "Not needed for sheets"
    Else
        ws.Range("C9").Value = "Metres of paper to use"
        If perJob Then
            ws.Range("C10").Value = "Required for student-supplied roll"
        ElseIf s.Found Then
            ws.Range("C10").Value = "Optional - blank uses the full " & Mm(s.WidthMM) & " mm"
        Else
            ws.Range("C10").Value = "Optional - blank uses the full roll width"
        End If
    End If
    If isSheet And perJob Then
        ws.Range("C11").Value = "Required - nearest standard size"
    Else
        ws.Range("C11").Value = "Only for student-supplied sheets"
    End If

    ' --- inputs that do not apply to this paper are greyed and emptied ----
    SetApplicable ws.Range(PLN_WIDTH), Not isSheet
    SetApplicable ws.Range(PLN_SIZE), (isSheet And perJob)

    ' --- estimate --------------------------------------------------------
    If Len(res.Message) = 0 Then
        ws.Range(PLN_PAPER_COST).Value = res.PaperCost
        ws.Range(PLN_INK_COST).Value = res.InkCost
        ws.Range(PLN_TOTAL_COST).Value = res.PaperCost + res.InkCost
        ws.Range(PLN_MESSAGE).Value = "Estimate at today's prices. Disregard flags are not applied, so this is the gross cost."
        ws.Range(PLN_MESSAGE).Font.Color = RGB(110, 110, 110)
        ws.Range(PLN_MESSAGE).Font.Bold = False
    Else
        ws.Range(PLN_PAPER_COST & ":" & PLN_TOTAL_COST).ClearContents
        ws.Range(PLN_MESSAGE).Value = res.Message
        If res.IsError Then
            ws.Range(PLN_MESSAGE).Font.Color = RGB(176, 0, 32)
            ws.Range(PLN_MESSAGE).Font.Bold = True
        Else
            ws.Range(PLN_MESSAGE).Font.Color = RGB(110, 110, 110)
            ws.Range(PLN_MESSAGE).Font.Bold = False
        End If
    End If
    ws.Range(PLN_MESSAGE).Font.Italic = True

    RelockSheet ws
    Application.EnableEvents = ev
End Sub

' Whole millimetres without a trailing point (VBA's "#,##0.##" gives "914.").
Private Function Mm(ByVal V As Double) As String
    If V = Int(V) Then
        Mm = Format$(V, "#,##0")
    Else
        Mm = Format$(V, "#,##0.0#")
    End If
End Function

Private Sub SetApplicable(ByVal target As Range, ByVal Applies As Boolean)
    If Applies Then
        target.Interior.Color = RGB(255, 255, 255)
    Else
        If Len(CStr(target.Value)) > 0 Then target.ClearContents
        target.Interior.Color = RGB(217, 217, 217)
    End If
End Sub

' ================================================================ events ===
' Workbook_SheetChange for the Summary sheet. Events are already off.
Public Sub PlannerChanged(ByVal ws As Worksheet, ByVal Target As Range)
    Dim planRange As Range, c As Range, picks As Range

    Set planRange = ws.Range("B7:B11,B19")
    If Intersect(Target, planRange) Is Nothing Then Exit Sub

    ' A marked "(unavailable)" pick is stored as the real name, as on a job row.
    Set picks = Intersect(Target, ws.Range("B7:B8"))
    If Not picks Is Nothing Then
        For Each c In picks.Cells
            If Len(CStr(c.Value)) > 0 Then
                If CleanPick(CStr(c.Value)) <> CStr(c.Value) Then c.Value = CleanPick(CStr(c.Value))
            End If
        Next c
    End If

    AppOff
    ' The other lists (and the marking of this one) follow the printer/paper pick.
    If Not picks Is Nothing Then BindPlannerLists ws
    PlannerRecalc ws
    AppOn
End Sub

' Re-reads prices and lists when the Summary is shown (a price may have been
' edited on Papers or Printers, or a print room added, since the box was last
' calculated). Cheap, and never touches the user's choices.
Public Sub RefreshPlanner(ByVal ws As Worksheet)
    On Error GoTo Fail
    If StrComp(ws.Name, "Summary", vbTextCompare) <> 0 Then Exit Sub
    AppOff
    BindPlannerLists ws
    PlannerRecalc ws
    AppOn
    Exit Sub
Fail:
    AppReset
End Sub

' Same, for callers that don't hold the sheet (RebindAllLocationDropdowns,
' the registry refresh). Quietly does nothing if there is no Summary sheet.
Public Sub RefreshSummaryPlanner()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Summary")
    On Error GoTo 0
    If Not ws Is Nothing Then RefreshPlanner ws
End Sub

' ================================================================== add ===
' The "Add to print room" button (modMain.btnPlannerAdd).
Public Sub AddPlannedJob(ByVal ws As Worksheet)
    Dim res As PlanResult, room As String, target As Worksheet, v As Variant, nm As String

    res = PlanFor(ws)
    If Len(res.Message) > 0 Then
        Say "The job is not ready to add yet.", res.Message, "Complete the plan above, then press Add to print room."
        Exit Sub
    End If

    room = Trim$(CStr(ws.Range(PLN_ROOM).Value))
    If Len(room) = 0 Then
        Say "No print room is chosen.", "The job needs a print room to be added to.", "Pick one from the Add to print room dropdown, then try again."
        Exit Sub
    End If

    For Each v In LocationSheets()
        nm = LocValue(v, "LOC_Name")
        If Len(nm) = 0 Then nm = v.Name
        If StrComp(nm, room, vbTextCompare) = 0 Then
            Set target = v
            Exit For
        End If
    Next v
    If target Is Nothing Then
        Say "'" & room & "' is not a print room.", "It may have been renamed or removed since it was chosen.", "Pick a room from the dropdown again."
        Exit Sub
    End If
    If Not InList(LocValue(target, "LOC_Printers"), res.Printer) Then
        Say res.Printer & " is not set up at " & room & ".", "A print room can only log jobs for the printers ticked in its Select printers... list.", "Choose a room that has this printer."
        Exit Sub
    End If

    CreatePlannedJob target, res
End Sub

Private Sub CreatePlannedJob(ByVal target As Worksheet, ByRef res As PlanResult)
    Dim lo As ListObject, r As ListRow, n As Long, s As clsStock, qty As Double

    Set lo = JobsTable(target)
    If lo Is Nothing Then
        Say "'" & target.Name & "' has no print job table.", "The sheet is marked as a print room but its job table is missing.", "Run Check workbook, or choose another room."
        Exit Sub
    End If
    Set s = Stock(res.Stock)

    qty = res.Qty
    If s.Measure <> "Sheet" Then
        If StrComp(RollUnitOf(target), "Centimetres", vbTextCompare) = 0 Then qty = qty * 100#
    End If

    AppOff
    UnlockSheet target
    If lo.ListRows.Count = 1 And IsBlankRow(lo, 1) Then
        Set r = lo.ListRows(1)
    Else
        Set r = lo.ListRows.Add
    End If
    n = r.Index

    CellIn(lo, n, "Job ID").Value = NextJobId(target, lo)
    CellIn(lo, n, "Date/Time").Value = Now
    CellIn(lo, n, "Printer").Value = res.Printer
    CellIn(lo, n, "Paper Stock").Value = res.Stock
    CellIn(lo, n, "Qty").Value = qty
    If s.Measure <> "Sheet" And res.WidthMM > 0 Then CellIn(lo, n, "Print Width mm").Value = res.WidthMM
    If s.Measure = "Sheet" And s.PerJobSize Then CellIn(lo, n, "Sheet size").Value = res.SizeName

    CellIn(lo, n, "Disregard Paper").Value = YesNoDefault(target, "LOC_DefDisPaper")
    CellIn(lo, n, "Disregard Consumable").Value = YesNoDefault(target, "LOC_DefDisCons")
    CellIn(lo, n, "Paid").Value = "No"
    If Len(LocValue(target, "LOC_DefTech")) > 0 Then CellIn(lo, n, "Technician").Value = LocValue(target, "LOC_DefTech")
    CellIn(lo, n, "Notes").Value = "Planned on the Summary page"

    RelockSheet target
    BindStockCell target, lo, n
    BindPrinterCell target, lo, n
    StampRow target, n
    AppOn

    ' Take the person to the new row so the student details can be typed in.
    ' Cosmetic, and impossible when unattended, so it must not fail the add.
    On Error Resume Next
    target.Activate
    CellIn(lo, n, "Student Name").Select
    On Error GoTo 0
End Sub

Private Function YesNoDefault(ByVal ws As Worksheet, ByVal RefName As String) As String
    If StrComp(LocValue(ws, RefName), "Yes", vbTextCompare) = 0 Then
        YesNoDefault = "Yes"
    Else
        YesNoDefault = "No"
    End If
End Function
