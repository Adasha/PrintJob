Attribute VB_Name = "modCatalog"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Configuration loaded once per operation.
'
' Reports and validation sweeps touch every job row; re-reading the Papers and
' Printers tables inside those loops is the classic way to make a workbook feel
' broken. Everything is read once into clsDict instances and looked up here.
'
' Stocks and printers are objects (clsStock, clsPrinterDef) rather than Public
' Types. VBA will not put a user-defined type declared in a standard module
' into a Variant or a Collection, and clsDict is a Collection underneath.
'
' A lookup that finds nothing returns an object with Found = False rather than
' Nothing, so callers can test .Found without guarding every reference.

' The two built-in student-supplied stocks (see "student-supplied stock" at
' the foot of this module). Reserved names: not tblPapers rows.
Public Const SUPPLIED_ROLL As String = "Supplied (Roll)"
Public Const SUPPLIED_SHEET As String = "Supplied (Sheet)"

Private mStocks As clsDict      ' by description
Private mPrinters As clsDict    ' by model
Private mTechs As clsDict       ' by name -> TechID
Private mLoaded As Boolean
Private mDirty As Boolean       ' set alongside mLoaded=False; cleared once
                                 ' every print room's dropdowns have actually
                                 ' been rebound (ThisWorkbook.Workbook_
                                 ' SheetDeactivate) - see modLists.
                                 ' RebindAllLocationDropdowns. Invalidate on
                                 ' its own only clears the in-memory cache;
                                 ' nothing pushes the change out to the
                                 ' dropdowns already drawn on each sheet until
                                 ' something rebinds them, and this flag is
                                 ' how that "something" knows it is needed.

Public Sub Invalidate()
    mLoaded = False
    mDirty = True
End Sub

' True from the moment a catalogue edit invalidates the cache until
' RebindAllLocationDropdowns next runs - i.e. "some print room's Printer/
' Paper Stock/Technician dropdown may now be showing stale data".
Public Function CatalogDirty() As Boolean
    CatalogDirty = mDirty
End Function

Public Sub ClearCatalogDirty()
    mDirty = False
End Sub

Public Sub LoadCatalog(Optional ByVal Force As Boolean = False)
    If mLoaded And Not Force Then Exit Sub

    Dim lo As ListObject, i As Long
    Dim s As clsStock, p As clsPrinterDef
    Dim szW As Double, szH As Double

    Set mStocks = New clsDict
    Set mPrinters = New clsDict
    Set mTechs = New clsDict

    Set lo = Tbl("tblPapers")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Description").Value) > 0 And Not IsBuiltInStock(CStr(CellIn(lo, i, "Description").Value)) Then
            Set s = New clsStock
            s.StockID = CStr(CellIn(lo, i, "StockID").Value)
            s.Description = CStr(CellIn(lo, i, "Description").Value)
            s.PaperType = CStr(CellIn(lo, i, "Paper type").Value)
            s.Measure = CStr(CellIn(lo, i, "Measure").Value)
            s.WidthMM = NumOf(CellIn(lo, i, "Width mm"))
            s.HeightMM = NumOf(CellIn(lo, i, "Height mm"))
            s.Cost = NumOf(CellIn(lo, i, "Cost"))
            s.Active = (StrComp(CStr(CellIn(lo, i, "Active").Value), "Yes", vbTextCompare) = 0)
            If ColumnExists(lo, "Supplied by student") Then
                s.SuppliedByStudent = (StrComp(CStr(CellIn(lo, i, "Supplied by student").Value), "Yes", vbTextCompare) = 0)
            End If
            ' Paper the student supplies is never charged for, whatever the
            ' Cost cell says - the flag decides, not the number.
            If s.SuppliedByStudent Then s.Cost = 0
            s.Found = True
            mStocks.Add s.Description, s
        End If
    Next i
    ' Last, so both stay at the bottom of the Paper Stock dropdown.
    AddBuiltInStocks

    Set lo = Tbl("tblPrinters")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Model").Value) > 0 Then
            Set p = New clsPrinterDef
            p.PrinterID = CStr(CellIn(lo, i, "PrinterID").Value)
            p.Model = CStr(CellIn(lo, i, "Model").Value)
            p.Consumable = CStr(CellIn(lo, i, "Consumable type").Value)
            p.RatePerM2 = NumOf(CellIn(lo, i, "Cost per m2"))
            If ColumnExists(lo, "Max roll width mm") Then
                p.MaxRollWidthMM = NumOf(CellIn(lo, i, "Max roll width mm"))
            End If
            If ColumnExists(lo, "Max sheet size") Then
                p.MaxSheetSize = Trim$(CStr(CellIn(lo, i, "Max sheet size").Value))
                If Len(p.MaxSheetSize) > 0 Then
                    ' Via local scratch variables, not p.MaxSheetWidthMM/
                    ' HeightMM directly - an object's public field passed as
                    ' a ByRef argument doesn't reliably write back over COM,
                    ' confirmed directly: StdSizeDims returned True with the
                    ' correct values, but p.MaxSheetWidthMM/HeightMM stayed 0
                    ' every time until routed through plain local Doubles.
                    If StdSizeDims(p.MaxSheetSize, szW, szH) Then
                        p.MaxSheetWidthMM = szW
                        p.MaxSheetHeightMM = szH
                    End If
                End If
            End If
            p.Active = (StrComp(CStr(CellIn(lo, i, "Active").Value), "Yes", vbTextCompare) = 0)
            p.Found = True
            mPrinters.Add p.Model, p
        End If
    Next i

    Set lo = Tbl("tblTechnicians")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Name").Value) > 0 Then
            mTechs.Add CStr(CellIn(lo, i, "Name").Value), CStr(CellIn(lo, i, "TechID").Value)
        End If
    Next i

    mLoaded = True
End Sub

Public Function Stock(ByVal Description As String) As clsStock
    LoadCatalog
    If mStocks.Exists(Description) Then
        Set Stock = mStocks.Obj(Description)
    Else
        Set Stock = New clsStock
    End If
End Function

Public Function Prn(ByVal Model As String) As clsPrinterDef
    LoadCatalog
    If mPrinters.Exists(Model) Then
        Set Prn = mPrinters.Obj(Model)
    Else
        Set Prn = New clsPrinterDef
    End If
End Function

Public Function TechID(ByVal TechName As String) As String
    LoadCatalog
    TechID = CStr(mTechs.Item(TechName))
End Function

' A stock is usable on a printer only when the printer can take that kind of
' stock (roll/sheet, via Max roll width mm / Max sheet size) AND the stock's
' own size fits within that capacity. (Compatibility used to be decided by a
' paper "family" standing in for a width band - "Short Roll"/"Long Roll" -
' which was too coarse; the family concept has since been removed entirely,
' leaving Papers' Measure column - Sheet or Roll - as the one input.)
'
' The two built-in "Supplied (Roll)"/"Supplied (Sheet)" stocks (student-
' supplied stock) have no fixed size to check here - the real dimensions
' aren't known until job entry (Print Width mm / the job-row Sheet size
' column), so compatibility for them is just "the printer takes this kind of
' stock at all". modValidation re-checks the row's own entered size against
' the chosen printer's capacity once it exists.
Public Function Compatible(ByVal Model As String, ByVal Description As String) As Boolean
    Dim p As clsPrinterDef, s As clsStock
    Set p = Prn(Model)
    Set s = Stock(Description)
    If Not p.Found Then Exit Function
    If Not s.Found Then Exit Function

    If s.Measure = "Roll" Then
        If p.MaxRollWidthMM <= 0 Then Exit Function
        If s.PerJobSize Then
            Compatible = True
        Else
            Compatible = (s.WidthMM <= p.MaxRollWidthMM)
        End If
    ElseIf s.Measure = "Sheet" Then
        If Len(p.MaxSheetSize) = 0 Then Exit Function
        If s.PerJobSize Then
            Compatible = True
        Else
            Compatible = FitsWithinMaxSheet(s.WidthMM, s.HeightMM, p.MaxSheetWidthMM, p.MaxSheetHeightMM)
        End If
    End If
End Function

' Whether a W x H stock fits within a MaxW x MaxH capacity in EITHER
' orientation - a smaller sheet can be fed either way round, so the
' comparison normalises both pairs to (longer side, shorter side) rather
' than assuming a fixed portrait/landscape match. Handles non-standard
' shapes (a Custom sheet stock with no Std. size name) exactly the same way
' as a named standard size, since it only ever looks at Width mm/Height mm.
Public Function FitsWithinMaxSheet(ByVal W As Double, ByVal H As Double, ByVal MaxW As Double, ByVal MaxH As Double) As Boolean
    If MaxW <= 0 Or MaxH <= 0 Then Exit Function
    If W <= 0 Or H <= 0 Then Exit Function
    FitsWithinMaxSheet = (Application.WorksheetFunction.Max(W, H) <= Application.WorksheetFunction.Max(MaxW, MaxH)) _
                      And (Application.WorksheetFunction.Min(W, H) <= Application.WorksheetFunction.Min(MaxW, MaxH))
End Function

' A human-readable description of what a printer can take, for the AT-05
' "incompatible combination" messages in modValidation - replaces the old
' "does not support the X paper family" phrasing now that compatibility is
' a size fit rather than a family list.
Public Function CapacityText(ByVal p As clsPrinterDef) As String
    Dim parts As String
    If p.MaxRollWidthMM > 0 Then parts = "roll stock up to " & Format$(p.MaxRollWidthMM, "#,##0") & " mm wide"
    If Len(p.MaxSheetSize) > 0 Then
        If Len(parts) > 0 Then parts = parts & " or "
        parts = parts & "sheet stock up to " & p.MaxSheetSize
    End If
    If Len(parts) = 0 Then parts = "no paper stock at all"
    CapacityText = parts
End Function

' Looks up a Standard Sizes row by name - shared by printer capacity
' resolution (LoadCatalog, above) and a student-supplied sheet job's own
' "Sheet size" cell (modSnapshot.StampRow, modValidation.OnSheetSizeChanged).
' Returns False (leaving W/H untouched) when the name isn't found, e.g. a
' size removed from the catalogue after being chosen.
Public Function StdSizeDims(ByVal SizeName As String, ByRef W As Double, ByRef H As Double) As Boolean
    Dim lo As ListObject, i As Long
    Set lo = Tbl("tblStandardSizes")
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Size name").Value)), Trim$(SizeName), vbTextCompare) = 0 Then
            W = NumOf(CellIn(lo, i, "Width mm"))
            H = NumOf(CellIn(lo, i, "Height mm"))
            StdSizeDims = True
            Exit Function
        End If
    Next i
End Function

' Active stocks compatible with the printer, for the dependent dropdown.
Public Function StocksFor(ByVal Model As String) As Collection
    Dim out As Collection, k As Variant, s As clsStock
    Set out = New Collection
    LoadCatalog
    For Each k In mStocks.Keys
        Set s = mStocks.Obj(CStr(k))
        If s.Active Then
            If Compatible(Model, s.Description) Then out.Add s.Description
        End If
    Next k
    Set StocksFor = out
End Function

' Active printers permitted at this location (spec 9.1, AT-06).
Public Function PrintersFor(ByVal ws As Worksheet) As Collection
    Dim out As Collection, k As Variant, p As clsPrinterDef, allowed As String
    Set out = New Collection
    LoadCatalog
    allowed = LocValue(ws, "LOC_Printers")
    For Each k In mPrinters.Keys
        Set p = mPrinters.Obj(CStr(k))
        If p.Active Then
            If InList(allowed, p.Model) Then out.Add p.Model
        End If
    Next k
    Set PrintersFor = out
End Function

' The reverse of PrintersFor/StocksFor: active printers permitted at this
' location whose capacity fits the given stock (spec 1a). Paper Stock can now be chosen before Printer, so Printer's own list
' has to be narrowable by the row's stock just as Stock's list is narrowable
' by the row's printer.
Public Function PrintersForStock(ByVal ws As Worksheet, ByVal Description As String) As Collection
    Dim out As Collection, k As Variant, p As clsPrinterDef, allowed As String
    Set out = New Collection
    LoadCatalog
    allowed = LocValue(ws, "LOC_Printers")
    For Each k In mPrinters.Keys
        Set p = mPrinters.Obj(CStr(k))
        If p.Active Then
            If InList(allowed, p.Model) Then
                If Compatible(p.Model, Description) Then out.Add p.Model
            End If
        End If
    Next k
    Set PrintersForStock = out
End Function

Public Function ActiveTechnicians() As Collection
    Dim out As Collection, lo As ListObject, i As Long
    Set out = New Collection
    Set lo = Tbl("tblTechnicians")
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(CellIn(lo, i, "Active").Value), "Yes", vbTextCompare) = 0 Then
            If Len(CellIn(lo, i, "Name").Value) > 0 Then
                out.Add CStr(CellIn(lo, i, "Name").Value)
            End If
        End If
    Next i
    Set ActiveTechnicians = out
End Function

' Every active printer/paper stock DEFINED in the catalogue, regardless of
' whether it has been used in a job yet - for the Reports page filters (snag
' list item 4), which deliberately do not restrict one by the other the way
' StocksFor does for job entry. Sorted, unlike PrintersFor/StocksFor, because
' nothing else about those two orders it for a caller and a filter dropdown
' is read by a person.
Public Function AllActivePrinters() As Collection
    Dim out As New Collection, k As Variant, p As clsPrinterDef
    LoadCatalog
    For Each k In mPrinters.Keys
        Set p = mPrinters.Obj(CStr(k))
        If p.Active Then out.Add p.Model
    Next k
    Set AllActivePrinters = SortedTextCollection(out)
End Function

Public Function AllActiveStocks() As Collection
    Dim out As New Collection, k As Variant, s As clsStock
    LoadCatalog
    For Each k In mStocks.Keys
        Set s = mStocks.Obj(CStr(k))
        If s.Active Then out.Add s.Description
    Next k
    Set AllActiveStocks = SortedTextCollection(out)
End Function

' ---------------------------------------------------- catalogue row edit ---
' Add/remove a row on one of the catalogue tables - Printers, Papers, Print
' Technicians (snag list item 17). Mirrors modJobs.AddPrintJob/RemoveRow:
' ListRows.Add keeps the new row's formatting and validation, because an
' Excel Table extends both automatically, and removing asks first and shows
' what is about to go - the same "never a bare are you sure" rule spec 10.12
' sets for job rows.
'
' Deleting a catalogue row never touches a job already recorded against it:
' every job snapshots the price it was costed at (the S_* columns), which is
' the whole point of the snapshot - see modExport's "content is the
' irreversible decision" note.
Public Sub AddCatalogRow(ByVal TableName As String)
    Dim lo As ListObject, r As ListRow, ws As Worksheet
    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent

    AppOff
    UnlockSheet ws
    If lo.ListRows.Count = 1 And IsBlankRow(lo, 1) Then
        Set r = lo.ListRows(1)
    Else
        Set r = lo.ListRows.Add
    End If
    ' ListRows.Add copies formatting and validation but no VALUES, so Active
    ' would be blank - and LoadCatalog reads only a literal "Yes" as active, so
    ' a row left blank is silently excluded from every dropdown.
    DefaultActive lo, r.Range.Cells(1, 1).Row - lo.DataBodyRange.Row + 1
    DefaultSupplied lo, r.Range.Cells(1, 1).Row - lo.DataBodyRange.Row + 1
    ' Technicians, printers and papers each get their site-prefixed ID here,
    ' so it is there the moment the row is - see "catalogue IDs" below.
    FillCatalogId lo, TableName, r.Index
    RelockSheet ws
    Invalidate
    AppOn

    r.Range.Cells(1, 1).Select
End Sub

' Sets Active to "Yes" on table row RowNo when the table has an Active column
' and the cell is blank. Never overwrites an existing value. Caller has already
' unlocked the sheet.
Private Sub DefaultActive(ByVal lo As ListObject, ByVal RowNo As Long)
    If Not ColumnExists(lo, "Active") Then Exit Sub
    If Len(Trim$(CStr(CellIn(lo, RowNo, "Active").Value))) = 0 Then
        CellIn(lo, RowNo, "Active").Value = "Yes"
    End If
End Sub

' Papers table only: a blank "Supplied by student" becomes "No" - a stock is
' charged for unless it is explicitly marked as student-supplied. Caller has
' already unlocked the sheet.
Private Sub DefaultSupplied(ByVal lo As ListObject, ByVal RowNo As Long)
    If Not ColumnExists(lo, "Supplied by student") Then Exit Sub
    If Len(Trim$(CStr(CellIn(lo, RowNo, "Supplied by student").Value))) = 0 Then
        CellIn(lo, RowNo, "Supplied by student").Value = "No"
    End If
End Sub

' A stock marked Supplied by student never has a paper cost (LoadCatalog
' enforces it); this keeps the visible Cost cell honest as well. Caller has
' already unlocked the sheet.
Private Sub ZeroSuppliedCost(ByVal lo As ListObject, ByVal RowNo As Long)
    If Not ColumnExists(lo, "Supplied by student") Then Exit Sub
    If StrComp(Trim$(CStr(CellIn(lo, RowNo, "Supplied by student").Value)), "Yes", vbTextCompare) <> 0 Then Exit Sub
    If NumOf(CellIn(lo, RowNo, "Cost")) = 0 Then Exit Sub
    CellIn(lo, RowNo, "Cost").Value = 0
    Say "Paper supplied by the student is not charged for.", _
        "Cost has been set to 0 on '" & Trim$(CStr(CellIn(lo, RowNo, "Description").Value)) & "' because 'Supplied by student' is Yes.", _
        "Set 'Supplied by student' to No if the print room buys and charges for this stock."
End Sub

' Called from ThisWorkbook.Workbook_SheetChange for every edit on the Papers
' sheet, with events already off (so the writes below do not re-enter it).
'
' 1. Size mode "Standard": fills Width mm / Height mm from the Std. size chosen,
'    as the note on the Papers sheet promises. Compatible() compares these two
'    numbers with the printer's capacity, so a sheet stock left at 0 x 0 fits
'    no printer and never reaches a dropdown. Fires only when Size mode or Std.
'    size was the cell edited, so a later hand-edit of Width/Height is kept.
' 2. Active: defaults to "Yes" on a row that has a Description but no Active
'    value - covers a row created by typing under the table (Excel auto-extend),
'    which never passes through AddCatalogRow.
Public Sub OnPaperEdited(ByVal Target As Range)
    Dim lo As ListObject, hit As Range, c As Range, ws As Worksheet
    Dim n As Long, hdr As String, sizeName As String
    Dim w As Double, h As Double, en As Long, ed As String

    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set hit = Application.Intersect(Target, lo.DataBodyRange)
    If hit Is Nothing Then Exit Sub
    Set ws = lo.Parent

    UnlockSheet ws
    On Error GoTo Fail
    For Each c In hit.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        hdr = CStr(lo.HeaderRowRange.Cells(1, c.Column - lo.Range.Column + 1).Value)

        If StrComp(hdr, "Size mode", vbTextCompare) = 0 Or StrComp(hdr, "Std. size", vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(CellIn(lo, n, "Size mode").Value)), "Standard", vbTextCompare) = 0 Then
                sizeName = Trim$(CStr(CellIn(lo, n, "Std. size").Value))
                If Len(sizeName) > 0 Then
                    ' Plain local Doubles, not fields of an object - see the
                    ' ByRef note in LoadCatalog.
                    If StdSizeDims(sizeName, w, h) Then
                        CellIn(lo, n, "Width mm").Value = w
                        CellIn(lo, n, "Height mm").Value = h
                    End If
                End If
            End If
        End If

        ' "Supplied (Roll)"/"Supplied (Sheet)" are built in (AddBuiltInStocks);
        ' a table row can not use either name.
        If IsBuiltInStock(CStr(CellIn(lo, n, "Description").Value)) Then
            Say "'" & Trim$(CStr(CellIn(lo, n, "Description").Value)) & "' is a built-in paper stock and can not be used as a name here.", _
                "The two 'Supplied' stocks are always available and are not stored on the Papers sheet.", _
                "Choose a different description. To record paper a student supplies in bulk, use your own name and set 'Supplied by student' to Yes."
            CellIn(lo, n, "Description").ClearContents
        End If

        If Len(Trim$(CStr(CellIn(lo, n, "Description").Value))) > 0 Then
            DefaultActive lo, n
            DefaultSupplied lo, n
            ZeroSuppliedCost lo, n
            FillCatalogId lo, "tblPapers", n
        End If
    Next c
    On Error GoTo 0
    RelockSheet ws
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    RelockSheet ws
    Err.Raise en, "OnPaperEdited", ed
End Sub

' ----------------------------------------------------- catalogue IDs ---
' Technicians, printers and papers each carry a synthetic ID (TechID,
' PrinterID, StockID) that nobody types by hand. It is allocated the way a
' Job ID is (modRegistry.NextJobId): the site (SET_SITE_ID), a table code and
' a four-digit number taken from a persisted high-water mark that only ever
' rises, e.g. MAIN-PRN-0001. The reason is the same too - two workbooks that
' each number from 1 would issue the same ID to different things, and a
' backup restore (modBackup.ApplyCatalogRows) matches rows by exactly this ID,
' so it would overwrite one site's printer with another's. With the site in
' the ID, rows from another workbook never collide with local ones and
' combining configurations appends instead of overwriting.
'
' Unlike a Job ID there is no per-location part - these tables are site-wide -
' so the counters live in tblSettings (TECH_ID_HWM, PRINTER_ID_HWM,
' STOCK_ID_HWM), not the per-location registry. They are marked "Read-only"
' in their Notes, which is what makes a restore skip them (a backup must
' never wind a counter back) and locks the cell against hand edits.
'
' As with NextJobId, a scan of the IDs currently in the table is only a FLOOR
' under the persisted value, never the source of truth: deleting the
' highest-numbered row lowers what the scan finds but never the counter.

' The four facts that differ between the three tables. False for any other
' table, so callers can pass a table name without checking it first.
Public Function CatalogIdSpec(ByVal TableName As String, ByRef IdHeader As String, ByRef Code As String, _
                              ByRef NameHeader As String, ByRef HwmKey As String) As Boolean
    Select Case TableName
        Case "tblTechnicians"
            IdHeader = "TechID": Code = "TCH": NameHeader = "Name": HwmKey = "TECH_ID_HWM"
        Case "tblPrinters"
            IdHeader = "PrinterID": Code = "PRN": NameHeader = "Model": HwmKey = "PRINTER_ID_HWM"
        Case "tblPapers"
            IdHeader = "StockID": Code = "STK": NameHeader = "Description": HwmKey = "STOCK_ID_HWM"
        Case Else
            Exit Function
    End Select
    CatalogIdSpec = True
End Function

' The three counter rows in tblSettings. Called by Setup, and again by
' SetCatalogHwm if a counter is ever found missing.
Public Sub EnsureCatalogIdSettings()
    EnsureSetting "TECH_ID_HWM", "Last technician ID number", "Read-only. The highest TechID number issued at this site. Only ever rises, so an ID is never reused."
    EnsureSetting "PRINTER_ID_HWM", "Last printer ID number", "Read-only. The highest PrinterID number issued at this site. Only ever rises, so an ID is never reused."
    EnsureSetting "STOCK_ID_HWM", "Last paper ID number", "Read-only. The highest StockID number issued at this site. Only ever rises, so an ID is never reused."
End Sub

' Sets a counter. Events are switched off around the write: it lands on the
' Settings sheet, whose Workbook_SheetChange would otherwise mark the whole
' catalogue dirty and rebind every dropdown for a change nobody made.
Private Sub SetCatalogHwm(ByVal HwmKey As String, ByVal Value As Long)
    Dim c As Range, ev As Boolean

    On Error Resume Next
    Set c = ThisWorkbook.Names("SET_" & HwmKey).RefersToRange
    On Error GoTo 0
    If c Is Nothing Then
        EnsureCatalogIdSettings
        On Error Resume Next
        Set c = ThisWorkbook.Names("SET_" & HwmKey).RefersToRange
        On Error GoTo 0
    End If
    If c Is Nothing Then Exit Sub

    ev = Application.EnableEvents
    Application.EnableEvents = False
    UnlockSheet c.Parent
    c.Value = Value
    RelockSheet c.Parent
    Application.EnableEvents = ev
End Sub

' The next ID for one catalogue table, advancing its counter.
Public Function NextCatalogId(ByVal lo As ListObject, ByVal TableName As String) As String
    Dim idHdr As String, code As String, nameHdr As String, hwmKey As String
    Dim prefix As String, hi As Long

    If Not CatalogIdSpec(TableName, idHdr, code, nameHdr, hwmKey) Then Exit Function
    prefix = SettingText("SITE_ID", "SITE") & "-" & code & "-"

    hi = CLng(SettingNum(hwmKey, 0))
    hi = CLng(Application.WorksheetFunction.Max(hi, ScanMaxSuffix(lo, idHdr, prefix)))
    hi = hi + 1
    SetCatalogHwm hwmKey, hi
    NextCatalogId = prefix & Format$(hi, "0000")
End Function

' Gives table row RowNo its ID if it has none; never touches an existing one.
' Caller has already unlocked the sheet.
Public Sub FillCatalogId(ByVal lo As ListObject, ByVal TableName As String, ByVal RowNo As Long)
    Dim idHdr As String, code As String, nameHdr As String, hwmKey As String

    If Not CatalogIdSpec(TableName, idHdr, code, nameHdr, hwmKey) Then Exit Sub
    If Not ColumnExists(lo, idHdr) Then Exit Sub
    If Len(Trim$(CStr(CellIn(lo, RowNo, idHdr).Value))) > 0 Then Exit Sub
    CellIn(lo, RowNo, idHdr).Value = NextCatalogId(lo, TableName)
End Sub

' Raises a counter to cover every ID now in the table under this site's
' prefix - for after rows have been written in bulk (a restore), where the
' incoming numbers may exceed anything allocated locally so far. Never lowers
' it.
Public Sub SyncCatalogHwm(ByVal lo As ListObject, ByVal TableName As String)
    Dim idHdr As String, code As String, nameHdr As String, hwmKey As String
    Dim cur As Long, seen As Long

    If Not CatalogIdSpec(TableName, idHdr, code, nameHdr, hwmKey) Then Exit Sub
    If Not ColumnExists(lo, idHdr) Then Exit Sub
    cur = CLng(SettingNum(hwmKey, 0))
    seen = ScanMaxSuffix(lo, idHdr, SettingText("SITE_ID", "SITE") & "-" & code & "-")
    If seen > cur Then SetCatalogHwm hwmKey, seen
End Sub

' Setup: gives every named row that has no ID one. Covers the sample rows the
' template ships with and any row added while events were off. Rows that
' already have an ID are left exactly as they are.
Public Sub EnsureCatalogIds()
    Dim tables As Variant, t As Variant, lo As ListObject, ws As Worksheet, i As Long
    Dim idHdr As String, code As String, nameHdr As String, hwmKey As String

    tables = Array("tblTechnicians", "tblPrinters", "tblPapers")
    For Each t In tables
        Set lo = Tbl(CStr(t))
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                If CatalogIdSpec(CStr(t), idHdr, code, nameHdr, hwmKey) Then
                    If ColumnExists(lo, idHdr) Then
                        Set ws = lo.Parent
                        UnlockSheet ws
                        For i = 1 To lo.ListRows.Count
                            If Len(Trim$(CStr(CellIn(lo, i, nameHdr).Value))) > 0 Then FillCatalogId lo, CStr(t), i
                        Next i
                        RelockSheet ws
                    End If
                End If
            End If
        End If
    Next t
End Sub

' Called from ThisWorkbook.Workbook_SheetChange for an edit on the Printers or
' Print Technicians sheet, with events already off (Papers has its own,
' OnPaperEdited). A row created by typing under the table never passes through
' AddCatalogRow, so this is what gives it an ID: once it has a name.
Public Sub OnCatalogEdited(ByVal TableName As String, ByVal Target As Range)
    Dim lo As ListObject, hit As Range, c As Range, ws As Worksheet, n As Long
    Dim idHdr As String, code As String, nameHdr As String, hwmKey As String
    Dim en As Long, ed As String, needsId As Boolean

    If Not CatalogIdSpec(TableName, idHdr, code, nameHdr, hwmKey) Then Exit Sub
    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set hit = Application.Intersect(Target, lo.DataBodyRange)
    If hit Is Nothing Then Exit Sub
    Set ws = lo.Parent

    ' Most edits need nothing from here (a price, a rename, a row deleted), and
    ' unlocking then re-locking the sheet for them is not free of side effects:
    ' it would re-protect a sheet the caller had deliberately left unlocked for
    ' a run of structural changes. So look first, and touch protection only when
    ' a named row really has no ID.
    For Each c In hit.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        If Len(Trim$(CStr(CellIn(lo, n, nameHdr).Value))) > 0 And _
           Len(Trim$(CStr(CellIn(lo, n, idHdr).Value))) = 0 Then
            needsId = True
            Exit For
        End If
    Next c
    If Not needsId Then Exit Sub

    UnlockSheet ws
    On Error GoTo Fail
    For Each c In hit.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        If Len(Trim$(CStr(CellIn(lo, n, nameHdr).Value))) > 0 Then FillCatalogId lo, TableName, n
    Next c
    On Error GoTo 0
    RelockSheet ws
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    RelockSheet ws
    Err.Raise en, "OnCatalogEdited", ed
End Sub

Public Sub RemoveCatalogRow(ByVal TableName As String)
    Dim lo As ListObject, ws As Worksheet, n As Long, c As Range
    Dim detail As String, i As Long

    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    If lo.DataBodyRange Is Nothing Then
        Say "There is nothing to remove.", "'" & TableName & "' has no rows."
        Exit Sub
    End If

    Set c = Application.Intersect(Selection.Cells(1, 1).EntireRow, lo.DataBodyRange)
    If c Is Nothing Then
        Say "No row is selected.", "This command works on the row the cursor is in.", "Click any cell in the row you want to remove, then try again."
        Exit Sub
    End If
    n = c.Row - lo.DataBodyRange.Row + 1

    For i = 1 To lo.ListColumns.Count
        If i > 1 Then detail = detail & "  "
        detail = detail & lo.ListColumns(i).Name & ": " & CStr(lo.DataBodyRange.Cells(n, i).Value)
    Next i

    If Not Ask("Remove this row?" & vbCrLf & vbCrLf & detail & vbCrLf & vbCrLf & _
        "Print jobs already recorded against it keep their frozen prices and are not affected. This cannot be undone.", _
        "Remove row") Then Exit Sub

    AppOff
    UnlockSheet ws
    lo.ListRows(n).Delete
    RelockSheet ws
    Invalidate
    AppOn
End Sub

' ------------------------------------------------------ catalogue clear ---
' "Clear table" on Print Technicians, Printers and Papers (direct user
' request, 2026-09-29): empties ONE catalogue table so a bureau can start that
' list afresh without touching the others - e.g. a new academic year with new
' technicians but the same printers. Only these three tables have the button;
' the Settings lookup tables do not.
'
' ClearCatalogTable asks first (naming the table, how many rows and which,
' and what the consequences are - the same "never a bare are you sure" rule
' RemoveCatalogRow and modJobs.ClearAll follow). DoClearCatalogTable is the
' unprompted work, split out because Ask always declines in quiet mode, so a
' test script could otherwise never exercise the clear itself.
'
' Recorded jobs are not touched: every job snapshots the price it was costed
' at (the S_* columns) - the same guarantee RemoveCatalogRow gives. What does
' change is every print room's dropdowns (Invalidate marks them stale; they
' rebind on the next sheet change).
Public Sub ClearCatalogTable(ByVal TableName As String)
    Dim lo As ListObject, n As Long, i As Long, label As String, keyHdr As String
    Dim listing As String, shown As Long, extra As String, msg As String

    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub

    ClearTableInfo TableName, label, keyHdr, extra
    n = RowCount(lo)
    If n = 0 Then
        Say "There is nothing to clear.", "The " & label & " table has no rows."
        Exit Sub
    End If

    ' Name the first few rows so the person can see WHAT is about to go,
    ' not just how many.
    For i = 1 To lo.ListRows.Count
        If Len(Trim$(CStr(CellIn(lo, i, keyHdr).Value))) > 0 Then
            If shown < 5 Then
                listing = listing & "   - " & Trim$(CStr(CellIn(lo, i, keyHdr).Value)) & vbCrLf
            End If
            shown = shown + 1
        End If
    Next i
    If shown > 5 Then listing = listing & "   ...and " & (shown - 5) & " more" & vbCrLf

    msg = "Clear table - " & label & vbCrLf & vbCrLf & _
        "This will permanently delete all " & n & " row" & IIf(n = 1, "", "s") & " from the " & label & " table:" & vbCrLf & _
        listing & vbCrLf & _
        "They will disappear from every print room's dropdowns straight away." & vbCrLf & _
        "Print jobs already recorded keep their frozen prices and are not affected." & vbCrLf & _
        extra & _
        "The other configuration sheets are not affected." & vbCrLf & vbCrLf & _
        "This cannot be undone - use Backup workbook first if you may want these back. Continue?"

    If Not Ask(msg, "Clear table") Then Exit Sub

    DoClearCatalogTable TableName
    Say n & " row" & IIf(n = 1, "", "s") & " deleted from the " & label & " table.", _
        "A note of what was removed has been kept in the workbook's audit log."
End Sub

' What the confirmation calls the table, which column names its rows, and any
' extra sentence that only applies to this table (ends with vbCrLf if set).
Private Sub ClearTableInfo(ByVal TableName As String, ByRef Label As String, ByRef KeyHdr As String, ByRef Extra As String)
    Select Case TableName
        Case "tblTechnicians"
            Label = "Print Technicians": KeyHdr = "Name"
        Case "tblPrinters"
            Label = "Printers": KeyHdr = "Model"
            Extra = "Each print room's 'Select printers' choice refers to printers by name, so rooms will need their printers re-selected once you add new ones." & vbCrLf
        Case "tblPapers"
            Label = "Papers": KeyHdr = "Description"
            Extra = "The built-in 'Supplied (Roll)' and 'Supplied (Sheet)' stocks are not stored here and stay available." & vbCrLf
        Case Else
            Err.Raise vbObjectError + 2101, "ClearCatalogTable", "Clear table is not available for '" & TableName & "'."
    End Select
End Sub

' The unprompted clear. Leaves exactly one blank row rather than zero rows:
' with DataBodyRange gone there is no row left for Excel to copy formatting
' and validation from, so a later ListRows.Add would produce a bare row - the
' same reason modJobs.ClearAll keeps row 1. AddCatalogRow already reuses that
' blank row (Count = 1 And IsBlankRow) and RowCount reports it as zero rows.
' Cells holding a formula are skipped rather than cleared, so a calculated
' column survives.
Public Sub DoClearCatalogTable(ByVal TableName As String)
    Dim lo As ListObject, ws As Worksheet, n As Long, i As Long, c As Range
    Dim label As String, keyHdr As String, extra As String

    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    ClearTableInfo TableName, label, keyHdr, extra
    Set ws = lo.Parent
    n = RowCount(lo)
    If n = 0 Then Exit Sub

    AppOff
    LogAudit "Clear table", label, n & " row" & IIf(n = 1, "", "s") & " deleted"
    UnlockSheet ws
    For i = lo.ListRows.Count To 2 Step -1
        lo.ListRows(i).Delete
    Next i
    For Each c In lo.ListRows(1).Range.Cells
        If Not c.HasFormula Then c.ClearContents
    Next c
    RelockSheet ws
    Invalidate
    AppOn
End Sub

' ---------------------------------------------------- standard sizes name ---
' A workbook-scoped name for tblStandardSizes[Size name], re-created (delete
' then re-add, same idiom modInit.EnsureLocName uses) rather than referenced
' directly as a structured reference in a cross-sheet Data Validation list:
' Validation.Add's Formula1 rejects a bare cross-sheet structured reference
' with a plain 1004 over COM automation, even though the same text works
' fine typed into the dialog by hand - a defined name is what makes a
' cross-sheet list source reliable here, same reasoning modLists.bas gives
' for why every OTHER dropdown in this project stages its list first rather
' than pointing straight at a table range.
Public Sub EnsureStdSizesName()
    On Error Resume Next
    ThisWorkbook.Names("RNG_STD_SIZES").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="RNG_STD_SIZES", RefersTo:="=tblStandardSizes[Size name]"
End Sub

' -------------------------------------------------- student-supplied stock -
' Two kinds of "supplied by the student" paper:
'
' 1. The built-in "Supplied (Roll)" / "Supplied (Sheet)" stocks. They are NOT
'    rows of tblPapers: they exist for every printer and location alike, so
'    LoadCatalog adds them in code (AddBuiltInStocks) and nobody can delete,
'    rename or re-cost them. The two names are reserved (IsBuiltInStock,
'    OnPaperEdited). They have no catalogue size - the real size is entered
'    on each job (Print Width mm / Sheet size) - and no paper cost.
' 2. Any tblPapers row with Supplied by student = Yes, e.g. a bulk delivery
'    of paper a student supplies for many jobs. It defaults to No, keeps its
'    own catalogue size like any other stock, and LoadCatalog forces its
'    Cost to 0.

Public Function IsBuiltInStock(ByVal Description As String) As Boolean
    Description = Trim$(Description)
    IsBuiltInStock = (StrComp(Description, SUPPLIED_ROLL, vbTextCompare) = 0) _
                  Or (StrComp(Description, SUPPLIED_SHEET, vbTextCompare) = 0)
End Function

Private Sub AddBuiltInStocks()
    AddBuiltInStock "STK-SUP-ROLL", SUPPLIED_ROLL, "Roll"
    AddBuiltInStock "STK-SUP-SHEET", SUPPLIED_SHEET, "Sheet"
End Sub

Private Sub AddBuiltInStock(ByVal StockID As String, ByVal Description As String, ByVal Measure As String)
    Dim s As New clsStock
    s.StockID = StockID
    s.Description = Description
    s.PaperType = "Student supplied"
    s.Measure = Measure             ' "Roll" / "Sheet" - fixed, not looked up
    s.Cost = 0
    s.Active = True
    s.SuppliedByStudent = True
    s.PerJobSize = True
    s.Found = True
    mStocks.Add s.Description, s
End Sub

' One-off migration, run by Setup (and harmless to repeat): earlier builds
' shipped the two Supplied stocks as ordinary tblPapers rows. They are built
' in now, so any such row is deleted. Jobs already recorded against them are
' unaffected - each carries its own frozen S_* snapshot.
Public Sub RemoveLegacySuppliedRows()
    Dim lo As ListObject, ws As Worksheet, i As Long, n As Long, id As String
    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set ws = lo.Parent

    For i = lo.ListRows.Count To 1 Step -1
        id = Trim$(CStr(CellIn(lo, i, "StockID").Value))
        If IsBuiltInStock(CStr(CellIn(lo, i, "Description").Value)) _
            Or StrComp(id, "STK-SUP-ROLL", vbTextCompare) = 0 _
            Or StrComp(id, "STK-SUP-SHEET", vbTextCompare) = 0 Then
            If n = 0 Then UnlockSheet ws
            lo.ListRows(i).Delete
            n = n + 1
        End If
    Next i

    If n > 0 Then
        RelockSheet ws
        Invalidate
        LogAudit "Migration", "Papers", n & " built-in 'Supplied' paper row" & IIf(n = 1, "", "s") & " removed from tblPapers (now built in)"
    End If
End Sub

' Setup: every stock defaults to "Supplied by student = No", and a stock that
' is Yes has a Cost of 0. Only fills blanks / zeroes a stray cost; never
' flips a Yes or No the user chose.
Public Sub NormaliseSuppliedFlags()
    Dim lo As ListObject, ws As Worksheet, i As Long, n As Long
    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not ColumnExists(lo, "Supplied by student") Then Exit Sub
    Set ws = lo.Parent

    UnlockSheet ws
    For i = 1 To lo.ListRows.Count
        If Len(Trim$(CStr(CellIn(lo, i, "Description").Value))) > 0 Then
            If Len(Trim$(CStr(CellIn(lo, i, "Supplied by student").Value))) = 0 Then
                CellIn(lo, i, "Supplied by student").Value = "No"
                n = n + 1
            End If
            If StrComp(Trim$(CStr(CellIn(lo, i, "Supplied by student").Value)), "Yes", vbTextCompare) = 0 Then
                If NumOf(CellIn(lo, i, "Cost")) <> 0 Then CellIn(lo, i, "Cost").Value = 0
            End If
        End If
    Next i
    RelockSheet ws
    If n > 0 Then Invalidate
End Sub

' The Yes/No dropdown and its help text on tblPapers[Supplied by student].
' Re-applied here so the wording matches the behaviour above rather than
' whatever the .xlsx template happens to carry.
Public Sub EnsureSuppliedColumnValidation()
    Dim lo As ListObject, ws As Worksheet
    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not ColumnExists(lo, "Supplied by student") Then Exit Sub
    Set ws = lo.Parent

    UnlockSheet ws
    With lo.ListColumns("Supplied by student").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Supplied by student"
        .InputMessage = "Yes for paper the student supplies, e.g. a bulk delivery used across many jobs: paper cost is always zero, whatever Cost says. No (the default) for stock the print room buys and charges for."
        .ErrorTitle = "Supplied by student"
        .ErrorMessage = "Choose Yes or No."
    End With
    RelockSheet ws
End Sub
